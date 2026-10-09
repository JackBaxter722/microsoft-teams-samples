import { mkdirSync, readFileSync, renameSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";

export const DECISIONS = ["keep", "delete-request", "reassign-request", "quarantine-request"] as const;
export type Decision = (typeof DECISIONS)[number];
export type LifecycleStage = "needs-review" | "candidate-unused" | "orphaned";

export type ReviewDecision = {
  decision: Decision;
  actorId: string;
  at: string;
  rationale?: string;
  reassigneeId?: string;
  reassigneeName?: string;
};

export type Review = {
  id: string;
  resourceId: string;
  resourceType: string;
  displayName: string;
  environmentId?: string;
  lifecycleStage: LifecycleStage;
  stageRationale: string;
  makerId: string;
  reviewerId: string;
  createdAt: string;
  status: "pending" | "maker-responded";
  decisions: ReviewDecision[];
};

type StoredData = {
  reviews: Review[];
  conversations: Record<string, string>;
};

const emptyData = (): StoredData => ({ reviews: [], conversations: {} });

export class ReviewStore {
  private data: StoredData;

  constructor(private readonly filename: string) {
    try {
      const parsed = JSON.parse(readFileSync(filename, "utf8")) as StoredData;
      this.data = {
        reviews: Array.isArray(parsed.reviews) ? parsed.reviews : [],
        conversations:
          parsed.conversations && typeof parsed.conversations === "object"
            ? parsed.conversations
            : {}
      };
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
      this.data = emptyData();
    }
  }

  private persist(): void {
    mkdirSync(dirname(this.filename), { recursive: true });
    const temporaryFile = `${this.filename}.tmp`;
    writeFileSync(temporaryFile, JSON.stringify(this.data, null, 2), { mode: 0o600 });
    renameSync(temporaryFile, this.filename);
  }

  all(): Review[] {
    return structuredClone(this.data.reviews);
  }

  forMaker(makerId: string): Review[] {
    return this.all().filter((review) => review.makerId === makerId);
  }

  create(input: Omit<Review, "id" | "createdAt" | "status" | "decisions">): Review {
    const review: Review = {
      ...input,
      id: crypto.randomUUID(),
      createdAt: new Date().toISOString(),
      status: "pending",
      decisions: []
    };
    this.data.reviews.push(review);
    this.persist();
    return structuredClone(review);
  }

  submitDecision(
    reviewId: string,
    actorId: string,
    decision: Decision,
    options: Pick<ReviewDecision, "rationale" | "reassigneeId" | "reassigneeName"> = {}
  ): Review {
    const review = this.data.reviews.find((item) => item.id === reviewId);
    if (!review) throw new Error("Review not found.");
    if (review.makerId !== actorId) throw new Error("Only the assigned maker can respond.");
    if (review.status !== "pending") throw new Error("This review has already been answered.");
    if (decision === "reassign-request" && !options.reassigneeId) {
      throw new Error("Choose a person before submitting a reassignment request.");
    }

    review.decisions.push({
      decision,
      actorId,
      at: new Date().toISOString(),
      ...options
    });
    review.status = "maker-responded";
    this.persist();
    return structuredClone(review);
  }

  saveConversation(userId: string, conversationId: string): void {
    this.data.conversations[userId] = conversationId;
    this.persist();
  }

  conversationFor(userId: string): string | undefined {
    return this.data.conversations[userId];
  }
}
