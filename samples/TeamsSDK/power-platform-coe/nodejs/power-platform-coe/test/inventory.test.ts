import assert from "node:assert/strict";
import { test } from "node:test";
import { queryInventory } from "../inventory.js";

test("inventory query follows API skip tokens and maps inventory rows", async () => {
  const bodies: Array<Record<string, any>> = [];
  const fakeFetch: typeof fetch = async (_input, init) => {
    bodies.push(JSON.parse(String(init?.body)));
    return new Response(
      JSON.stringify({
        totalRecords: 101,
        skipToken: bodies.length === 1 ? "next-page-token" : "",
        data: [
          {
            name: "/environments/env-1/apps/app-1",
            type: "microsoft.powerapps/canvasapps",
            properties: {
              displayName: "Inventory app",
              ownerId: "maker-1",
              environmentId: "env-1",
              createdAt: "2025-01-01T00:00:00Z"
            }
          }
        ]
      }),
      { status: 200, headers: { "Content-Type": "application/json" } }
    );
  };

  const first = await queryInventory("delegated-token", undefined, undefined, fakeFetch);
  const second = await queryInventory("delegated-token", first.skipToken, undefined, fakeFetch);
  assert.equal(first.resources[0].displayName, "Inventory app");
  assert.equal(first.resources[0].ownerId, "maker-1");
  assert.equal(bodies[0].TableName, "PowerPlatformResources");
  assert.equal(bodies[0].Options.Skip, 0);
  assert.equal(bodies[1].Options.SkipToken, "next-page-token");
  assert.equal("Skip" in bodies[1].Options, false);
  assert.equal(second.totalRecords, 101);
});

test("resource IDs are escaped in the exact-resource inventory filter", async () => {
  let submitted: Record<string, any> | undefined;
  const fakeFetch: typeof fetch = async (_input, init) => {
    submitted = JSON.parse(String(init?.body));
    return new Response(JSON.stringify({ data: [], totalRecords: 0 }), { status: 200 });
  };
  await queryInventory("delegated-token", undefined, "resource'with-quote", fakeFetch);
  assert.equal(submitted?.Clauses[0].Values[0], "'resource''with-quote'");
});
