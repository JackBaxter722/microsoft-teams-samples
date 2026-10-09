# Power Platform COE lifecycle review

This sample combines a Teams bot and a personal tab to help a Power Platform Center of Excellence (COE) team request and track maker reviews of inventoried resources.

The sample uses the current Teams SDK, Fluent UI React components, Microsoft Teams SSO, and the Power Platform Inventory API. Inventory access is delegated to the signed-in user. That user must have the documented `ResourceQuery.Resources.Read` permission and an Entra role supported by the Inventory API. The sample does not use an app-only token.

> **Important limitations:** The Inventory API is a query API, not a resource-management API. This sample records review decisions only; it does not delete, block, quarantine, or transfer resources. The API schema does not provide a general last-used signal, so “unused” or “orphaned” are administrator-assessed review labels, not conclusions inferred from age or last-modified time. Inventory results can take up to 15 minutes to reflect changes, and flow owner metadata may represent the creator rather than the current owner.

## Features

- Administrators query resources visible to their delegated account, choose a maker, and send a proactive review card.
- Makers can request a card of pending reviews from the bot, respond with Keep, Delete request, Reassign request, or Quarantine request, and query their visible resources in the personal tab.
- The tab shows inventory results, lifecycle review labels, decisions and history. Search uses a Fluent UI combobox; reassignment uses Teams' people picker and asks for confirmation before recording the request.
- All review submissions are revalidated by the server and appended to a local sample data file. This file store is for local demonstration only; use a secured, durable database and retention policy in production.

## Inventory API and identity setup

1. Create an Entra app registration for the Teams app and configure the Teams SSO resource/scope `access_as_user`. Set `TAB_AUDIENCE` to that API audience, typically `api://<tab-domain>/<client-id>`.
2. Add the delegated Power Platform API permission `ResourceQuery.Resources.Read` and grant consent as required by your tenant.
3. Assign the signed-in administrators an Entra role supported by Inventory API. The API also applies its own tenant-role visibility restrictions. Set `ADMIN_USER_IDS` to the administrators' Entra object IDs for this sample's additional review-creation check.
4. Register/provision the Teams bot with the Teams Developer CLI and configure `CLIENT_ID`, `CLIENT_SECRET`, and `TENANT_ID`. The bot credentials are used only by the Teams SDK to send bot messages; inventory calls use on-behalf-of (OBO) tokens for the signed-in tab user.
5. Copy `.env.TEMPLATE` to `.env` and set its values. Never commit `.env`.

The backend obtains a Power Platform token through OBO for `https://api.powerplatform.com/.default` and calls:

`POST https://api.powerplatform.com/resourcequery/resources/query?api-version=2024-10-01`

It queries the `PowerPlatformResources` table and follows the API's `skipToken` paging field. The tab filters maker-visible results to the signed-in user's object ID; administrators can view the inventory records their delegated account is permitted to query. If the API does not expose an owner ID for a resource type, it is not included in a maker's list.

## Run locally

Requirements: Node.js 20+, a Microsoft 365 developer tenant, Teams Developer CLI, and a public HTTPS tunnel.

```bash
cd samples/TeamsSDK/power-platform-coe/nodejs/power-platform-coe
npm install
cp .env.TEMPLATE .env
# Configure the values described above.
npm run build
npm start
```

The app listens on port 3978. Provision/package the Teams app with the Teams Developer CLI and upload it to the tenant. Add the app to a personal chat, then open its **Power Platform COE** tab. Configure your tunnel and Teams app SSO resource consistently; the tab and API share the bot's HTTPS origin.

In the bot chat:

- `resources` opens the tab for an inventory query.
- `reviews` returns your pending maker review cards.
- `help` shows the available commands.

Administrators can create review cards from the tab. Assign stages and rationale based on a separate usage/ownership assessment; Inventory API creation/modification metadata alone is not proof that a resource is unused or orphaned.

## Validation and production notes

Run `npm test` for the review-state and inventory paging tests, and `npm run build` for the TypeScript and tab bundles. Before production use, replace the local JSON store, define and enforce an approval policy for Delete/Reassign/Quarantine requests, add audit/retention controls, and verify API permissions and lifecycle action APIs independently.

## References

- [Power Platform Inventory API](https://learn.microsoft.com/en-us/power-platform/admin/inventory-api)
- [Inventory schema](https://learn.microsoft.com/en-us/power-platform/admin/inventory-schema)
- [Inventory overview, roles, latency, and limitations](https://learn.microsoft.com/en-us/power-platform/admin/power-platform-inventory)
- [Power Platform programmability permissions](https://learn.microsoft.com/en-us/power-platform/admin/programmability-permission-reference)
- [Power Platform programmability authentication](https://learn.microsoft.com/en-us/power-platform/admin/programmability-authentication-v2)
