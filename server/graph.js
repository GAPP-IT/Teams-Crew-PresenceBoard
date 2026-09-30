import { ClientSecretCredential } from "@azure/identity";

const credential = new ClientSecretCredential(
  process.env.TENANT_ID,
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET
);

export async function getAccessToken() {
  const token = await credential.getToken(
    "https://graph.microsoft.com/.default"
  );

  return token.token;