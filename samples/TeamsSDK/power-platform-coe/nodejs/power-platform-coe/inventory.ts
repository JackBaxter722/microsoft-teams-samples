const INVENTORY_URL =
  "https://api.powerplatform.com/resourcequery/resources/query?api-version=2024-10-01";
const PAGE_SIZE = 100;

export type InventoryResource = {
  id: string;
  type: string;
  displayName: string;
  environmentId?: string;
  ownerId?: string;
  createdAt?: string;
  lastModifiedAt?: string;
};

export type InventoryPage = {
  resources: InventoryResource[];
  totalRecords: number;
  skipToken?: string;
};

type InventoryRow = {
  name?: unknown;
  type?: unknown;
  properties?: {
    displayName?: unknown;
    environmentId?: unknown;
    ownerId?: unknown;
    createdAt?: unknown;
    lastModifiedAt?: unknown;
  };
};

type QueryResponse = {
  data?: InventoryRow[];
  totalRecords?: number;
  skipToken?: string;
};

function stringValue(value: unknown): string | undefined {
  return typeof value === "string" && value.length > 0 ? value : undefined;
}

function mapResource(row: InventoryRow): InventoryResource | undefined {
  const id = stringValue(row.name);
  const type = stringValue(row.type);
  if (!id || !type) return undefined;

  const properties = row.properties ?? {};
  return {
    id,
    type,
    displayName: stringValue(properties.displayName) ?? id,
    environmentId: stringValue(properties.environmentId),
    ownerId: stringValue(properties.ownerId),
    createdAt: stringValue(properties.createdAt),
    lastModifiedAt: stringValue(properties.lastModifiedAt)
  };
}

export async function queryInventory(
  accessToken: string,
  skipToken?: string,
  resourceId?: string,
  request: typeof fetch = fetch
): Promise<InventoryPage> {
  const options: { Top: number; Skip?: number; SkipToken?: string } = {
    Top: PAGE_SIZE
  };
  if (skipToken) options.SkipToken = skipToken;
  else {
    options.Skip = 0;
    options.SkipToken = "";
  }

  const clauses: Array<Record<string, unknown>> = [];
  if (resourceId) {
    clauses.push({
      $type: "where",
      FieldName: "name",
      Operator: "==",
      Values: [`'${resourceId.replaceAll("'", "''")}'`]
    });
  }
  clauses.push({
    $type: "project",
    FieldList: [
      "name",
      "type",
      "properties.displayName",
      "properties.environmentId",
      "properties.ownerId",
      "properties.createdAt",
      "properties.lastModifiedAt"
    ]
  });

  const response = await request(INVENTORY_URL, {
    method: "POST",
    headers: {
      Authorization: `******
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      TableName: "PowerPlatformResources",
      Options: options,
      Clauses: clauses
    })
  });

  if (!response.ok) {
    throw new Error(`Power Platform Inventory API returned HTTP ${response.status}.`);
  }

  const result = (await response.json()) as QueryResponse;
  return {
    resources: (result.data ?? [])
      .map(mapResource)
      .filter((resource): resource is InventoryResource => resource !== undefined),
    totalRecords: Number.isFinite(result.totalRecords) ? result.totalRecords! : 0,
    skipToken: stringValue(result.skipToken)
  };
}
