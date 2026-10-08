import {
  app,
  authentication,
} from "@microsoft/teams-js";

let teamsInitialized = false;

async function ensureTeamsInitialized() {
  if (teamsInitialized) {
    return;
  }

  await app.initialize();

  teamsInitialized = true;
}

export async function apiFetch(
  url,
  options = {},
) {
  await ensureTeamsInitialized();

  const token =
    await authentication.getAuthToken();

  if (!token) {
    throw new Error(
      "No Teams authentication token received.",
    );
  }

  const headers =
    new Headers(options.headers || {});

  headers.set(
    "Authorization",
    `Bearer ${token}`,
  );

  if (
    options.body &&
    !headers.has("Content-Type")
  ) {
    headers.set(
      "Content-Type",
      "application/json",
    );
  }

  return fetch(url, {
    ...options,
    headers,
  });
}