import { ConfidentialClientApplication } from "@azure/msal-node";
import { createRemoteJWKSet, jwtVerify, type JWTPayload } from "jose";

export type UserClaims = JWTPayload & {
  oid: string;
  tid: string;
  scp: string;
};

export class AuthenticationError extends Error {}

let cachedTenantId: string | undefined;
let cachedKeys: ReturnType<typeof createRemoteJWKSet> | undefined;

function required(name: string): string {
  const value = process.env[name];
  if (!value) throw new Error(`Missing required configuration: ${name}`);
  return value;
}

function getClient(): ConfidentialClientApplication {
  return new ConfidentialClientApplication({
    auth: {
      clientId: required("CLIENT_ID"),
      authority: `https://login.microsoftonline.com/${required("TENANT_ID")}`,
      clientSecret: required("CLIENT_SECRET")
    }
  });
}

export async function verifyUserToken(authorization?: string): Promise<UserClaims> {
  const token = authorization?.startsWith("Bearer ") ? authorization.slice(7) : "";
  if (!token) throw new AuthenticationError("A Teams SSO bearer token is required.");
  const tenantId = required("TENANT_ID");
  const issuers = [
    `https://login.microsoftonline.com/${tenantId}/v2.0`,
    `https://sts.windows.net/${tenantId}/`
  ];
  if (cachedTenantId !== tenantId || !cachedKeys) {
    cachedTenantId = tenantId;
    cachedKeys = createRemoteJWKSet(
      new URL(`https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`)
    );
  }

  try {
    const { payload } = await jwtVerify(token, cachedKeys, {
      issuer: issuers,
      audience: required("TAB_AUDIENCE"),
      clockTolerance: 5
    });
    const claims = payload as UserClaims;
    if (
      claims.tid !== tenantId ||
      typeof claims.oid !== "string" ||
      !claims.scp?.split(" ").includes("access_as_user")
    ) {
      throw new AuthenticationError("The token does not grant access to this tab API.");
    }
    return claims;
  } catch (error) {
    if (error instanceof AuthenticationError) throw error;
    throw new AuthenticationError("The Teams SSO token is invalid or expired.");
  }
}

export async function getPowerPlatformToken(userAssertion: string): Promise<string> {
  const result = await getClient().acquireTokenOnBehalfOf({
    oboAssertion: userAssertion,
    scopes: [process.env.POWER_PLATFORM_SCOPE ?? "https://api.powerplatform.com/.default"]
  });
  if (!result?.accessToken) {
    throw new Error("Could not acquire a delegated Power Platform token.");
  }
  return result.accessToken;
}

export function isAdministrator(userId: string): boolean {
  const configuredIds = (process.env.ADMIN_USER_IDS ?? "")
    .split(",")
    .map((id) => id.trim().toLowerCase())
    .filter(Boolean);
  return configuredIds.includes(userId.toLowerCase());
}
