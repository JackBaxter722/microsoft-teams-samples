import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { ReviewStore } from "../reviews.js";

test("review decisions are maker-scoped, one-time, and persisted", () => {
  const directory = mkdtempSync(join(tmpdir(), "coe-reviews-"));
  try {
    const filename = join(directory, "reviews.json");
    const store = new ReviewStore(filename);
    const review = store.create({
      resourceId: "/resource/1",
      resourceType: "app",
      displayName: "Sample app",
      lifecycleStage: "candidate-unused",
      stageRationale: "Usage review needed",
      makerId: "maker-1",
      reviewerId: "admin-1"
    });

    assert.throws(
      () => store.submitDecision(review.id, "maker-2", "keep"),
      /Only the assigned maker/
    );
    assert.throws(
      () => store.submitDecision(review.id, "maker-1", "reassign-request"),
      /Choose a person/
    );

    const answered = store.submitDecision(review.id, "maker-1", "reassign-request", {
      reassigneeId: "maker-3",
      reassigneeName: "New owner"
    });
    assert.equal(answered.status, "maker-responded");
    assert.throws(
      () => store.submitDecision(review.id, "maker-1", "keep"),
      /already been answered/
    );

    const reloaded = new ReviewStore(filename);
    assert.equal(reloaded.forMaker("maker-1")[0].decisions[0].reassigneeId, "maker-3");
    assert.deepEqual(reloaded.forMaker("maker-2"), []);
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
});
