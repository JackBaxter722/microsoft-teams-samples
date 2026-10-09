import {
  AdaptiveCard,
  ExecuteAction,
  TextBlock
} from "@microsoft/teams.cards";
import type { Review } from "./reviews.js";

export function createReviewCard(review: Review): AdaptiveCard {
  const warning =
    "This card records your response for review. It does not change, delete, transfer, or quarantine the Power Platform resource.";
  return new AdaptiveCard(
    new TextBlock(`**Power Platform resource review: ${review.displayName}**`),
    new TextBlock(
      `Stage: ${review.lifecycleStage}\nResource type: ${review.resourceType}\n` +
        `Review context: ${review.stageRationale}\n\n${warning}`
    )
  ).withActions(
    new ExecuteAction({
      title: "Keep",
      verb: "record_decision",
      data: { reviewId: review.id, decision: "keep" }
    }),
    new ExecuteAction({
      title: "Request delete",
      verb: "record_decision",
      data: { reviewId: review.id, decision: "delete-request" }
    }),
    new ExecuteAction({
      title: "Reassign in tab",
      verb: "record_decision",
      data: { reviewId: review.id, decision: "reassign-request" }
    }),
    new ExecuteAction({
      title: "Request quarantine",
      verb: "record_decision",
      data: { reviewId: review.id, decision: "quarantine-request" }
    })
  );
}
