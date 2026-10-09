import { App } from "@microsoft/teams.apps";
import type { AdaptiveCardActionMessageResponse, IMessageActivity } from "@microsoft/teams.api";
import { fileURLToPath } from "node:url";
import { isAdministrator, getPowerPlatformToken, verifyUserToken, AuthenticationError } from "./auth.js";
import { createReviewCard } from "./cards.js";
import { queryInventory } from "./inventory.js";
import { DECISIONS, ReviewStore, type Decision, type LifecycleStage } from "./reviews.js";

const app = new App();
const http = app.http;
if (!http) throw new Error("The Teams SDK HTTP plugin is not configured.");
const dataFile = process.env.DATA_FILE ?? "./data/reviews.json";
const store = new ReviewStore(dataFile);
const webDirectory = fileURLToPath(new URL("./web/dist", import.meta.url));
const lifecycleStages: LifecycleStage[] = ["needs-review", "candidate-unused", "orphaned"];

app.tab("power-platform-coe", webDirectory);

function sendFailure(res: { status: (status: number) => { json: (body: object) => void } }, error: unknown) {
  if (error instanceof AuthenticationError) {
    res.status(401).json({ error: error.message });
  } else if (error instanceof Error && error.message === "Forbidden") {
    res.status(403).json({ error: "Administrator permissions are required." });
  } else {
    console.error("[COE] Request failed:", error);
    res.status(500).json({ error: "The request could not be completed." });
  }
}

function bearerFrom(authorization?: string): string {
  if (!authorization?.startsWith("Bearer ")) {
    throw new AuthenticationError("A Teams SSO bearer token is required.");
  }
  return authorization.slice(7);
}

http.get("/api/session", async (req, res) => {
  try {
    const user = await verifyUserToken(req.headers.authorization);
    res.json({ userId: user.oid, isAdministrator: isAdministrator(user.oid) });
  } catch (error) {
    sendFailure(res, error);
  }
});

http.get("/api/resources", async (req, res) => {
  try {
    const user = await verifyUserToken(req.headers.authorization);
    const accessToken = await getPowerPlatformToken(bearerFrom(req.headers.authorization));
    const skipToken =
      typeof req.query.skipToken === "string" ? req.query.skipToken : undefined;
    const page = await queryInventory(accessToken, skipToken);
    const resources = isAdministrator(user.oid)
      ? page.resources
      : page.resources.filter((resource) => resource.ownerId?.toLowerCase() === user.oid.toLowerCase());
    res.json({ ...page, resources });
  } catch (error) {
    sendFailure(res, error);
  }
});

http.get("/api/reviews", async (req, res) => {
  try {
    const user = await verifyUserToken(req.headers.authorization);
    const reviews = isAdministrator(user.oid) ? store.all() : store.forMaker(user.oid);
    res.json({ reviews });
  } catch (error) {
    sendFailure(res, error);
  }
});

http.post("/api/reviews", async (req, res) => {
  try {
    const user = await verifyUserToken(req.headers.authorization);
    if (!isAdministrator(user.oid)) throw new Error("Forbidden");

    const { resourceId, makerId, makerName, lifecycleStage, rationale } = req.body ?? {};
    if (
      typeof resourceId !== "string" ||
      !resourceId ||
      typeof makerId !== "string" ||
      !makerId ||
      typeof lifecycleStage !== "string" ||
      !lifecycleStages.includes(lifecycleStage as LifecycleStage) ||
      typeof rationale !== "string" ||
      !rationale.trim()
    ) {
      res.status(400).json({ error: "Choose a resource, maker, review stage, and rationale." });
      return;
    }

    const accessToken = await getPowerPlatformToken(bearerFrom(req.headers.authorization));
    const result = await queryInventory(accessToken, undefined, resourceId);
    const resource = result.resources.find((item) => item.id === resourceId);
    if (!resource) {
      res.status(404).json({ error: "The resource was not found in the current inventory query." });
      return;
    }

    const review = store.create({
      resourceId: resource.id,
      resourceType: resource.type,
      displayName: resource.displayName,
      environmentId: resource.environmentId,
      lifecycleStage: lifecycleStage as LifecycleStage,
      stageRationale: rationale.trim().slice(0, 500),
      makerId,
      reviewerId: user.oid
    });

    const conversationId = store.conversationFor(makerId);
    if (!conversationId) {
      res.status(201).json({
        review,
        notification: "Review saved. The maker must open the bot once before proactive cards can be delivered."
      });
      return;
    }
    try {
      await app.send(conversationId, createReviewCard(review));
      res.status(201).json({ review, notification: `Review card sent to ${makerName || "maker"}.` });
    } catch (error) {
      console.error("[COE] Review saved but proactive delivery failed:", error);
      res.status(201).json({ review, notification: "Review saved, but the bot could not deliver its card." });
    }
  } catch (error) {
    sendFailure(res, error);
  }
});

http.post("/api/reviews/:reviewId/decision", async (req, res) => {
  try {
    const user = await verifyUserToken(req.headers.authorization);
    const { decision, rationale, reassigneeId, reassigneeName } = req.body ?? {};
    if (!DECISIONS.includes(decision as Decision)) {
      res.status(400).json({ error: "Choose a supported review decision." });
      return;
    }
    if (
      decision === "reassign-request" &&
      (typeof reassigneeId !== "string" || !reassigneeId || reassigneeId === user.oid)
    ) {
      res.status(400).json({ error: "Select a different person before requesting reassignment." });
      return;
    }

    const review = store.submitDecision(req.params.reviewId, user.oid, decision as Decision, {
      rationale: typeof rationale === "string" ? rationale.trim().slice(0, 500) : undefined,
      reassigneeId: typeof reassigneeId === "string" ? reassigneeId : undefined,
      reassigneeName: typeof reassigneeName === "string" ? reassigneeName.slice(0, 200) : undefined
    });
    res.json({ review });
  } catch (error) {
    if (error instanceof Error && error.message === "Review not found.") {
      res.status(404).json({ error: error.message });
    } else if (
      error instanceof Error &&
      (error.message.startsWith("Only the assigned maker") ||
        error.message.startsWith("This review has") ||
        error.message.startsWith("Choose a person"))
    ) {
      res.status(409).json({ error: error.message });
    } else {
      sendFailure(res, error);
    }
  }
});

app.on("message", async ({ activity, send }) => {
  const userId = activity.from.aadObjectId;
  if (!userId || !activity.conversation.id) {
    await send("I could not identify this Teams user. Open the Power Platform COE tab and sign in.");
    return;
  }
  store.saveConversation(userId, activity.conversation.id);
  const text = ((activity as IMessageActivity).text ?? "").trim().toLowerCase();

  if (text.includes("reviews")) {
    const pending = store.forMaker(userId).filter((review) => review.status === "pending");
    if (pending.length === 0) {
      await send("You have no pending Power Platform resource reviews.");
      return;
    }
    await send(`You have ${pending.length} pending review(s).`);
    for (const review of pending) await send(createReviewCard(review));
    return;
  }
  if (text.includes("resources")) {
    await send("Open the Power Platform COE personal tab to query the inventory using your signed-in account.");
    return;
  }
  await send(
    "Power Platform COE review bot. Send `resources` to open the inventory tab or `reviews` to receive your pending review cards."
  );
});

app.on("card.action", async ({ activity, send }) => {
  const userId = activity.from.aadObjectId;
  const actionData = (activity as any).value?.action?.data;
  if (actionData?.decision === "reassign-request") {
    await send("Open the Power Platform COE tab to select a proposed new owner and confirm your reassignment request.");
    return {
      statusCode: 200,
      type: "application/vnd.microsoft.activity.message",
      value: "Continue in the Power Platform COE tab."
    } satisfies AdaptiveCardActionMessageResponse;
  }
  if (!userId || actionData?.decision !== "keep" && actionData?.decision !== "delete-request" &&
      actionData?.decision !== "reassign-request" && actionData?.decision !== "quarantine-request") {
    await send("This review action is invalid. Open the Power Platform COE tab to review it.");
    return {
      statusCode: 200,
      type: "application/vnd.microsoft.activity.message",
      value: "Invalid review action."
    } satisfies AdaptiveCardActionMessageResponse;
  }

  try {
    const updated = store.submitDecision(
      actionData.reviewId,
      userId,
      actionData.decision as Decision
    );
    await send(
      `Your "${actionData.decision}" response was recorded for ${updated.displayName}. ` +
        "This records a review request only; no resource lifecycle operation was performed."
    );
  } catch {
    await send("This review is no longer pending or is assigned to another maker.");
  }
  return {
    statusCode: 200,
    type: "application/vnd.microsoft.activity.message",
    value: "Review response recorded."
  } satisfies AdaptiveCardActionMessageResponse;
});

const port = Number(process.env.PORT ?? 3978);
app.start(port).catch((error) => {
  console.error("[COE] Teams app failed to start:", error);
  process.exitCode = 1;
});
