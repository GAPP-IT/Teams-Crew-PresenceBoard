import { mkdir, readFile, rename, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { createRemoteJWKSet, jwtVerify } from "jose";

const tenantId = process.env.TENANT_ID;
const ssoResource = process.env.TEAMS_SSO_RESOURCE;
const appUrl = process.env.TEAMS_APP_URL;
const storePath = resolve(process.env.TEAMS_NOTIFICATIONS_STORE_PATH || "data/teams-notifications.json");
const pollIntervalMs = Math.max(15000, Number(process.env.TEAMS_NOTIFICATION_POLL_INTERVAL_MS) || 30000);
const configured = Boolean(tenantId && process.env.CLIENT_ID && process.env.CLIENT_SECRET && ssoResource && appUrl?.startsWith("https://"));
const signingKeys = tenantId ? createRemoteJWKSet(new URL(`https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`)) : null;

let storedState = { subscriptions: {} };
let loadPromise;
let writeQueue = Promise.resolve();
let pollTimer;
let polling = false;
let previousStatuses = new Map();
let lastPollAt = null;
let lastPollError = null;
let lastNotificationSentAt = null;
let lastNotificationError = null;

async function loadStore() {
  if (!loadPromise) {
    loadPromise = readFile(storePath, "utf8")
      .then((content) => {
        const parsed = JSON.parse(content);
        storedState = {
          subscriptions: parsed?.subscriptions && typeof parsed.subscriptions === "object" ? parsed.subscriptions : {},
        };
      })
      .catch((error) => {
        if (error.code !== "ENOENT") throw error;
      });
  }
  await loadPromise;
}

function persistStore() {
  const content = JSON.stringify(storedState, null, 2);
  writeQueue = writeQueue
    .catch(() => {})
    .then(async () => {
      await mkdir(dirname(storePath), { recursive: true });
      const temporaryPath = `${storePath}.${process.pid}.tmp`;
      await writeFile(temporaryPath, content, { encoding: "utf8", mode: 0o600 });
      await rename(temporaryPath, storePath);
    });
  return writeQueue;
}

function serviceError(message, status) {
  return Object.assign(new Error(message), { status });
}

async function resolveTeamsUser(authorization = "") {
  if (!configured) throw serviceError("Teams notifications are not configured on the server.", 503);
  const assertion = /^Bearer\s+(.+)$/i.exec(authorization)?.[1];
  if (!assertion) throw serviceError("A Teams SSO token is required.", 401);

  try {
    const { payload } = await jwtVerify(assertion, signingKeys, {
      audience: ssoResource,
      issuer: [`https://login.microsoftonline.com/${tenantId}/v2.0`, `https://sts.windows.net/${tenantId}/`],
      algorithms: ["RS256"],
      clockTolerance: 5,
    });
    if (payload.tid !== tenantId || typeof payload.oid !== "string") {
      throw new Error("Missing Teams user identity.");
    }
    return payload.oid;
  } catch {
    throw serviceError("The Teams SSO token is invalid or expired.", 401);
  }
}

async function subscriptionsFor(authorization) {
  const userId = await resolveTeamsUser(authorization);
  await loadStore();
  return storedState.subscriptions[userId] || [];
}

async function replaceSubscriptions(authorization, personIds, getMembers) {
  const userId = await resolveTeamsUser(authorization);
  if (!Array.isArray(personIds) || personIds.length > 500) {
    throw serviceError("Invalid notification subscription list.", 400);
  }
  const uniqueIds = [...new Set(personIds.map(String))];
  if (uniqueIds.some((id) => !/^[\da-f-]{36}$/i.test(id))) {
    throw serviceError("Invalid person identifier in notification subscription.", 400);
  }

  const { members } = await getMembers();
  const allowedIds = new Set(members.map((member) => String(member.id)));
  if (uniqueIds.some((id) => !allowedIds.has(id))) {
    throw serviceError("A notification target is not part of the Presence Board.", 400);
  }

  await loadStore();
  if (uniqueIds.length) storedState.subscriptions[userId] = uniqueIds;
  else delete storedState.subscriptions[userId];
  await persistStore();
  return uniqueIds.length;
}

function notificationStatus(person) {
  if (person.presence === "away") return "away";
  if (person.phone === "call") return "call";
  if (["available", "busy", "meeting"].includes(person.presence)) return "free";
  return "other";
}

async function removeSubscription(userId, personId) {
  await loadStore();
  const currentIds = storedState.subscriptions[userId] || [];
  const nextIds = currentIds.filter((id) => id !== personId);
  if (nextIds.length) storedState.subscriptions[userId] = nextIds;
  else delete storedState.subscriptions[userId];
  await persistStore();
}

function errorStatus(error) {
  return /Graph (\d{3})/.exec(error.message)?.[1] || "send-failed";
}

export function createTeamsNotificationService({ getMembers, presences, mapPresence, graph }) {
  async function sendNotification(recipientId, person) {
    await graph(`/users/${encodeURIComponent(recipientId)}/teamwork/sendActivityNotification`, {
      method: "POST",
      body: JSON.stringify({
        topic: {
          source: "text",
          value: person.name,
          webUrl: appUrl,
        },
        activityType: "personAvailableAgain",
        previewText: { content: `${person.name} ist wieder verfügbar.` },
        templateParameters: [{ name: "person", value: person.name }],
      }),
    });
  }

  async function poll() {
    if (!configured || polling) return;
    polling = true;
    try {
      const { members } = await getMembers();
      const presencesById = new Map((await presences(members.map((member) => member.id))).map((presence) => [String(presence.id), presence]));
      const people = members.map((member) => ({
        id: String(member.id),
        name: member.displayName || "",
        ...mapPresence(presencesById.get(String(member.id))),
      }));
      const nextStatuses = new Map(people.map((person) => [person.id, notificationStatus(person)]));

      for (const person of people) {
        const previousStatus = previousStatuses.get(person.id);
        const nextStatus = nextStatuses.get(person.id);
        if (!previousStatus || !["call", "away"].includes(previousStatus) || nextStatus !== "free") continue;

        await loadStore();
        const recipients = Object.entries(storedState.subscriptions)
          .filter(([, personIds]) => personIds.includes(person.id))
          .map(([userId]) => userId);

        for (const recipientId of recipients) {
          try {
            await sendNotification(recipientId, person);
            await removeSubscription(recipientId, person.id);
            lastNotificationSentAt = new Date().toISOString();
            lastNotificationError = null;
          } catch (error) {
            lastNotificationError = errorStatus(error);
            console.error(`[teams-notifications] Graph send failed (${lastNotificationError}).`);
          }
        }
      }

      previousStatuses = nextStatuses;
      lastPollAt = new Date().toISOString();
      lastPollError = null;
    } catch (error) {
      lastPollError = errorStatus(error);
      console.error(`[teams-notifications] Poll failed (${lastPollError}).`);
    } finally {
      polling = false;
    }
  }

  return {
    getSubscriptions: subscriptionsFor,
    replaceSubscriptions: (authorization, personIds) => replaceSubscriptions(authorization, personIds, getMembers),
    start() {
      if (!configured || pollTimer) return;
      void poll();
      pollTimer = setInterval(() => void poll(), pollIntervalMs);
      pollTimer.unref?.();
    },
    health() {
      const subscriptions = Object.values(storedState.subscriptions);
      return {
        configured,
        activeUsers: subscriptions.filter((personIds) => personIds.length > 0).length,
        activeSubscriptions: subscriptions.reduce((count, personIds) => count + personIds.length, 0),
        pollIntervalMs,
        lastPollAt,
        lastPollError,
        lastNotificationSentAt,
        lastNotificationError,
      };
    },
  };
}
