import "dotenv/config";
import express from "express";
import cors from "cors";
import { ClientSecretCredential } from "@azure/identity";
import { createTeamsNotificationService } from "./teams-notifications.js";
import { mapCalendar, mapAutomaticReplies, mapPresence } from "./utils/mapping.js";

import helmet from "helmet";
import rateLimit from "express-rate-limit";
import { z } from "zod";
import {
  createRemoteJWKSet,
  jwtVerify,
} from "jose";

const app = express();
const PORT = Number(process.env.PORT || 3051);
const GRAPH = process.env.GRAPH_BASE_URL || "https://graph.microsoft.com/v1.0";
const GROUP_NAME = process.env.PRESENCE_GROUP_NAME || "ShowInPresenceBoard";

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;

const subscriptionsSchema = z.object({
  personIds: z
    .array(z.string().uuid())
    .max(650),
});

const userIdSchema = z.string().uuid();

/*
 * Must exactly match:
 *
 * Entra ID
 *   Expose an API
 *   Application ID URI
 *
 * and:
 *
 * Teams manifest
 *   webApplicationInfo.resource
 */
const API_AUDIENCE =
  process.env.API_AUDIENCE ||
  `api://${CLIENT_ID}`;

const REQUIRED_API_SCOPE =
  process.env.API_SCOPE ||
  "access_as_user";

if (!TENANT_ID) {
  throw new Error(
    "TENANT_ID environment variable is missing.",
  );
}

if (!CLIENT_ID) {
  throw new Error(
    "CLIENT_ID environment variable is missing.",
  );
}

if (!process.env.CLIENT_SECRET) {
  throw new Error(
    "CLIENT_SECRET environment variable is missing.",
  );
}

const validIssuers = [
  `https://login.microsoftonline.com/${TENANT_ID}/v2.0`,
  `https://sts.windows.net/${TENANT_ID}/`,
];

const entraJwks = createRemoteJWKSet(
  new URL(
    `https://login.microsoftonline.com/` +
      `${TENANT_ID}/discovery/v2.0/keys`,
  ),
);


const MEETING_ROOM_COMPANY = "__meeting_rooms__";
const MISSING_COMPANY = "__missing_company__";
const MISSING_DEPARTMENT = "__missing_department__";
const credential = new ClientSecretCredential(process.env.TENANT_ID, process.env.CLIENT_ID, process.env.CLIENT_SECRET);
const teamsNotifications = createTeamsNotificationService({ getMembers, presences, mapPresence, graph });

const allowedOrigins = new Set(
  (process.env.CORS_ALLOWED_ORIGINS || "")
    .split(",")
    .map((origin) => origin.trim())
    .filter(Boolean),
);

app.disable("x-powered-by");

/*
 * Set this when Apache is the controlled HTTPS
 * reverse proxy in front of Node.
 */
app.set("trust proxy", 1);

app.use(
  helmet({
    /*
     * The frontend runs inside a Teams iframe.
     * Frame embedding is normally controlled at
     * Apache/CSP level for the Teams domains.
     */
    frameguard: false,
  }),
);

app.use(
  cors({
    origin(origin, callback) {
      /*
       * Requests without Origin include monitoring
       * tools and server-to-server requests.
       */
      if (
        !origin ||
        allowedOrigins.has(origin)
      ) {
        return callback(null, true);
      }

      return callback(
        new Error("Origin not allowed by CORS."),
      );
    },

    methods: ["GET", "PUT", "OPTIONS"],

    allowedHeaders: [
      "Authorization",
      "Content-Type",
    ],

    credentials: false,
    maxAge: 86400,
  }),
);

app.use(
  express.json({
    limit: "32kb",
    strict: true,
  }),
);

const apiRateLimiter = rateLimit({
  windowMs: 60_000,
  limit: 120,
  standardHeaders: "draft-8",
  legacyHeaders: false,

  message: {
    success: false,
    error: "Too many requests.",
  },
});

app.use("/api", apiRateLimiter);

let memberCache = { expires: 0, group: null, members: [] };

function readBearerToken(request) {
  const authorization =
    request.get("authorization");

  if (!authorization) {
    return null;
  }

  const match =
    /^Bearer\s+(.+)$/i.exec(authorization);

  return match?.[1]?.trim() || null;
}

function readScopes(payload) {
  if (typeof payload.scp !== "string") {
    return [];
  }

  return payload.scp
    .split(" ")
    .map((scope) => scope.trim())
    .filter(Boolean);
}

async function requireTeamsUser(
  request,
  response,
  next,
) {
  try {
    const accessToken =
      readBearerToken(request);

    if (!accessToken) {
      return response.status(401).json({
        success: false,
        error: "Authentication required.",
      });
    }

    /*
     * jwtVerify validates:
     *
     * - cryptographic signature
     * - expiration
     * - not-before
     * - audience
     *
     * The issuer is checked separately so both supported
     * tenant-specific issuer formats can be handled.
     */
    const { payload, protectedHeader } =
      await jwtVerify(
        accessToken,
        entraJwks,
        {
          audience: API_AUDIENCE,
          algorithms: ["RS256"],
        },
      );

    if (
      typeof payload.iss !== "string" ||
      !validIssuers.includes(payload.iss)
    ) {
      return response.status(401).json({
        success: false,
        error: "Invalid token issuer.",
      });
    }

    if (payload.tid !== TENANT_ID) {
      return response.status(403).json({
        success: false,
        error:
          "Access from this tenant is not allowed.",
      });
    }

    if (payload.ver !== "2.0") {
      return response.status(401).json({
        success: false,
        error:
          "Unsupported access-token version.",
      });
    }

    const scopes = readScopes(payload);

    if (!scopes.includes(REQUIRED_API_SCOPE)) {
      return response.status(403).json({
        success: false,
        error:
          "Required API permission is missing.",
      });
    }

    if (
      typeof payload.oid !== "string" ||
      !payload.oid
    ) {
      return response.status(401).json({
        success: false,
        error:
          "Authenticated user identifier is missing.",
      });
    }

    request.user = {
      id: payload.oid,
      tenantId: payload.tid,
      name:
        typeof payload.name === "string"
          ? payload.name
          : null,
      username:
        typeof payload.preferred_username ===
        "string"
          ? payload.preferred_username
          : null,
      scopes,
      tokenId:
        typeof payload.uti === "string"
          ? payload.uti
          : null,
      keyId: protectedHeader.kid || null,
    };

    return next();
  } catch (error) {
    console.warn(
      "[authentication] Token validation failed:",
      error instanceof Error
        ? error.message
        : "Unknown authentication error",
    );

    return response.status(401).json({
      success: false,
      error:
        "Invalid or expired access token.",
    });
  }
}

async function token() {
  const result =
    await credential.getToken(GRAPH_SCOPE);

  if (!result?.token) {
    throw new Error("Kein Graph-Token erhalten.");
  }

  return result.token;
}

async function graph(path, options = {}) {
  const response = await fetch(path.startsWith("http") ? path : `${GRAPH}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${await token()}`,
      Accept: "application/json",
      ...(options.body ? { "Content-Type": "application/json" } : {}),
      ...options.headers,
    },
  });
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Graph ${response.status}: ${text}`);
  }
  if (response.status === 204) return null;
  return response.json();
}

async function collection(path, headers = {}) {
  const all = [];
  let next = path;
  while (next) {
    const d = await graph(next, { headers });
    all.push(...(d.value || []));
    next = d["@odata.nextLink"] || null;
  }
  return all;
}

async function getMembers() {
  if (Date.now() < memberCache.expires) return memberCache;
  const escaped = GROUP_NAME.replaceAll("'", "''");
  const groups = await graph(`/groups?$filter=${encodeURIComponent(`displayName eq '${escaped}'`)}&$select=id,displayName`);
  if ((groups.value || []).length !== 1) throw new Error(`Gruppe ${GROUP_NAME} wurde nicht eindeutig gefunden.`);
  const group = groups.value[0];
  const members = (
    await collection(
      `/groups/${group.id}/transitiveMembers/microsoft.graph.user?$select=id,displayName,givenName,surname,mail,userPrincipalName,jobTitle,department,companyName,officeLocation,accountEnabled`,
    )
  ).filter((u) => u.accountEnabled !== false);
  memberCache = { expires: Date.now() + 86400000, group, members };
  return memberCache;
}

function chunks(a, n) {
  const out = [];
  for (let i = 0; i < a.length; i += n) out.push(a.slice(i, i + n));
  return out;
}

async function presences(ids) {
  const out = [];
  for (const batch of chunks(ids, 650)) {
    const d = await graph("/communications/getPresencesByUserId", { method: "POST", body: JSON.stringify({ ids: batch }) });
    out.push(...(d.value || []));
  }
  return out;
}

async function calendar(userId, rangeDays = 1) {
  const now = new Date(),
    start = new Date(now),
    end = new Date(now);
  start.setDate(start.getDate() - 1);
  end.setDate(end.getDate() + rangeDays);
  const path = `/users/${userId}/calendar/calendarView?startDateTime=${encodeURIComponent(start.toISOString())}&endDateTime=${encodeURIComponent(end.toISOString())}&$select=id,start,end,showAs,isAllDay,isCancelled&$orderby=start/dateTime`;
  return collection(path, { Prefer: 'outlook.timezone="Europe/Berlin"' });
}

async function mailboxSettings(userId) {
  return graph(`/users/${encodeURIComponent(userId)}/mailboxSettings/automaticRepliesSetting`, { headers: { Prefer: 'outlook.timezone="UTC"' } });
}

async function roomPlaces() {
  return collection("/places/microsoft.graph.room?$select=emailAddress,capacity,building,floorNumber,floorLabel,label");
}

function initials(u) {
  if (u.givenName || u.surname) return `${u.givenName?.[0] || ""}${u.surname?.[0] || ""}`.toUpperCase();
  return (u.displayName || "")
    .replace(",", " ")
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((x) => x[0])
    .join("")
    .toUpperCase();
}
function company(name = "") {
  const s = String(name || "")
    .trim()
    .toLowerCase();
  if (s.includes("meeting room")) return MEETING_ROOM_COMPANY;
  if (!s) return MISSING_COMPANY;
  if (s.includes("global aviation") || s.includes("gapp")) return "GAPP";
  if (s.includes("european aviation") || s.includes("eac2") || s.includes("eacc")) return "EAC2";
  return "PAG";
}
async function build() {
  const { group, members } = await getMembers();
  const ps = await presences(members.map((x) => x.id));
  const byId = new Map(ps.map((x) => [x.id, x]));
  const roomMembers = members.filter((u) => company(u.companyName) === MEETING_ROOM_COMPANY);
  const calendarResults = await Promise.allSettled(
    members.map(async (u) => [u.id, await calendar(u.id, company(u.companyName) === MEETING_ROOM_COMPANY ? 30 : 1)]),
  );
  const calendars = new Map(calendarResults.filter((x) => x.status === "fulfilled").map((x) => x.value));
  const roomIds = new Set(roomMembers.map((u) => u.id));
  const failedRoomCalendars = calendarResults.filter((result, index) => result.status === "rejected" && roomIds.has(members[index].id));
  if (failedRoomCalendars.length)
    console.warn(
      `[presence-board] Calendar.Read failed for ${failedRoomCalendars.length}/${roomMembers.length} rooms: ${failedRoomCalendars[0].reason?.message || "unknown Graph error"}`,
    );
  const places = new Map();
  if (roomMembers.length) {
    try {
      const results = await roomPlaces();
      for (const place of results) {
        if (place.emailAddress) places.set(place.emailAddress.toLowerCase(), place);
      }
    } catch (error) {
      console.warn(`[presence-board] Places.Read failed: ${error.message || "unknown Graph error"}`);
    }
  }
  const mailboxResults = await Promise.allSettled(members.map(async (u) => [u.id, await mailboxSettings(u.id)]));
  const mailboxFailures = mailboxResults.flatMap((result, index) =>
    result.status === "rejected" ? [{ user: members[index], error: result.reason }] : [],
  );
  if (mailboxFailures.length) {
    const details = mailboxFailures
      .map(({ user, error }) => {
        const mailbox = user.mail || user.userPrincipalName;
        const identity = user.displayName || mailbox || user.id;
        return `${identity}${mailbox && mailbox !== identity ? ` (${mailbox})` : ""}: ${error?.message || "unknown Graph error"}`;
      })
      .join("; ");
    console.warn(`[presence-board] MailboxSettings.Read failed for ${mailboxFailures.length}/${members.length} users: ${details}`);
  }
  const mailboxes = new Map(mailboxResults.filter((x) => x.status === "fulfilled").map((x) => x.value));
  const people = members.map((u) => {
    const p = mapPresence(byId.get(u.id)),
      c = mapCalendar(calendars.get(u.id) || []),
      sageAbsence = null,
      companyKey = company(u.companyName),
      place = places.get((u.mail || u.userPrincipalName || "").toLowerCase());
    const resourceLocationDescription = [u.officeLocation, place?.building, place?.floorLabel || place?.floorNumber, place?.label]
      .filter(Boolean)
      .map(String)
      .filter((value, index, values) => values.indexOf(value) === index)
      .join(" · ");
    const automaticReplies = mapAutomaticReplies(mailboxes.get(u.id));
    const automaticRepliesAbsence = automaticReplies?.active ? { returnDate: automaticReplies.returnDate, note: automaticReplies.note } : null;
    const oofFromFallback = sageAbsence
      ? null
      : c.oofAbsence
        ? { ...c.oofAbsence, note: automaticReplies?.note || null }
        : p.outOfOffice
          ? { returnDate: null, note: automaticReplies?.note || null }
          : null;
    const oofAbsence = sageAbsence ? null : automaticReplies?.configured ? automaticRepliesAbsence : oofFromFallback;
    const absent = Boolean(sageAbsence || oofAbsence);
    return {
      id: u.id,
      name: u.displayName || "",
      initials: initials(u),
      email: u.mail || u.userPrincipalName || "",
      department: u.department || MISSING_DEPARTMENT,
      company: companyKey,
      companyName: u.companyName || "",
      role: u.jobTitle || "",
      officeLocation: u.officeLocation || "",
      resourceCapacity: companyKey === MEETING_ROOM_COMPANY ? (place?.capacity ?? null) : undefined,
      resourceLocationDescription: companyKey === MEETING_ROOM_COMPANY ? resourceLocationDescription || null : undefined,
      photoUrl: `/api/users/${u.id}/photo`,
      location: p.location,
      presence: p.presence,
      availability: p.availability,
      activity: p.activity,
      outOfOffice: p.outOfOffice,
      outOfOfficeNote: automaticReplies?.note || null,
      automaticRepliesStatus: automaticReplies?.status || "unavailable",
      automaticRepliesActive: Boolean(automaticReplies?.active),
      phone: absent ? "free" : p.phone,
      currentMeetingEnd: absent ? null : c.currentMeetingEnd,
      currentMeetingShowAs: absent ? null : c.currentMeetingShowAs,
      meetingReachable: absent ? false : c.meetingReachable,
      nextMeeting: absent || c.currentMeetingEnd ? null : c.nextMeeting,
      resourceSchedule: companyKey === MEETING_ROOM_COMPANY ? c.schedule : undefined,
      resourceCalendarAvailable: companyKey === MEETING_ROOM_COMPANY ? calendars.has(u.id) : undefined,
      sageAbsence,
      oofAbsence,
    };
  });
  return { success: true, generatedAt: new Date().toISOString(), group, count: people.length, people };
}



app.get(
  "/api/health",
  (request, response) => {
    response.json({
      success: true,
      status: "ok",
      generatedAt:
        new Date().toISOString(),
    });
  },
);

app.get(
  "/api/presence-board",
  requireTeamsUser,
  async (request, response) => {
    try {
      response.set(
        "Cache-Control",
        "no-store",
      );

      console.info(
        "[presence-board] Request",
        {
          userId: request.user.id,
          username: request.user.username,
          tenantId: request.user.tenantId,
        },
      );

      response.json(await build());
    } catch (error) {
      console.error(
        "[presence-board] Build failed:",
        error,
      );

      response.status(500).json({
        success: false,
        error:
          "Presence-Daten konnten nicht geladen werden.",
      });
    }
  },
);

/*app.get("/api/presence-board", async (req, res) => {
  try {
    res.set("Cache-Control", "no-store");
    res.json(await build());
  } catch (e) {
    console.error(e);
    res.status(500).json({ success: false, error: "Presence-Daten konnten nicht geladen werden.", details: e.message });
  }
});*/

app.get("/api/teams/notifications/subscriptions", async (req, res) => {
  try {
    res.set("Cache-Control", "no-store");
    res.json({ personIds: await teamsNotifications.getSubscriptions(req.get("authorization")) });
  } catch (error) {
    res.status(error.status || 500).json({ error: error.status ? error.message : "Teams subscriptions could not be loaded." });
  }
});



app.put(
  "/api/teams/notifications/subscriptions",
  requireTeamsUser,
  async (request, response) => {
    try {
      const parsed =
        subscriptionsSchema.safeParse(
          request.body,
        );

      if (!parsed.success) {
        return response.status(400).json({
          success: false,
          error:
            "Invalid subscription request.",
        });
      }

      const activeSubscriptions =
        await teamsNotifications
          .replaceSubscriptions(
            request.get("authorization"),
            parsed.data.personIds,
          );

      return response.json({
        success: true,
        activeSubscriptions,
      });
    } catch (error) {
      console.error(
        "[subscriptions] Saving failed:",
        error,
      );

      const status =
        Number.isInteger(error?.status) &&
        error.status >= 400 &&
        error.status < 500
          ? error.status
          : 500;

      return response.status(status).json({
        success: false,
        error:
          status === 500
            ? "Teams subscriptions could not be saved."
            : error.message,
      });
    }
  },
);

/*app.put("/api/teams/notifications/subscriptions", async (req, res) => {
  try {
    const activeSubscriptions = await teamsNotifications.replaceSubscriptions(req.get("authorization"), req.body?.personIds);
    res.json({ success: true, activeSubscriptions });
  } catch (error) {
    res.status(error.status || 500).json({ error: error.status ? error.message : "Teams subscriptions could not be saved." });
  }
});*/

app.get(
  "/api/users/:userId/photo",
  requireTeamsUser,
  async (request, response) => {
    const parsedUserId =
      userIdSchema.safeParse(
        request.params.userId,
      );

    if (!parsedUserId.success) {
      return response.status(400).json({
        success: false,
        error: "Invalid user ID.",
      });
    }

    try {
      /*
       * Prevent authenticated users from using this
       * endpoint as an unrestricted tenant photo proxy.
       */
      const { members } = await getMembers();

      const requestedUser =
        members.find(
          (member) =>
            member.id === parsedUserId.data,
        );

      if (!requestedUser) {
        return response.sendStatus(404);
      }

      const graphResponse = await fetch(
        `${GRAPH}/users/` +
          `${encodeURIComponent(parsedUserId.data)}` +
          "/photo/$value",
        {
          signal:
            AbortSignal.timeout(15000),

          headers: {
            Authorization:
              `Bearer ${await token()}`,
          },
        },
      );

      if (graphResponse.status === 404) {
        return response.sendStatus(404);
      }

      if (!graphResponse.ok) {
        console.error(
          "[photo] Graph request failed:",
          graphResponse.status,
        );

        return response.status(502).json({
          success: false,
          error:
            "Profile photo could not be loaded.",
        });
      }

      response.set(
        "Content-Type",
        graphResponse.headers.get(
          "content-type",
        ) || "image/jpeg",
      );

      response.set(
        "Cache-Control",
        "private, max-age=86400",
      );

      return response.send(
        Buffer.from(
          await graphResponse.arrayBuffer(),
        ),
      );
    } catch (error) {
      console.error(
        "[photo] Request failed:",
        error,
      );

      return response.status(502).json({
        success: false,
        error:
          "Profile photo could not be loaded.",
      });
    }
  },
);

/*app.get("/api/users/:userId/photo", async (req, res) => {
  try {
    const response = await fetch(`${GRAPH}/users/${encodeURIComponent(req.params.userId)}/photo/$value`, {
      headers: { Authorization: `Bearer ${await token()}` },
    });
    if (!response.ok) return res.sendStatus(404);
    res.set("Content-Type", response.headers.get("content-type") || "image/jpeg");
    res.set("Cache-Control", "private, max-age=86400");
    res.send(Buffer.from(await response.arrayBuffer()));
  } catch {
    return res.sendStatus(404);
  }
});*/

app.use((req, res) => res.status(404).json({ success: false, error: "API-Endpunkt nicht gefunden.", path: req.originalUrl }));

app.listen(PORT, () => console.log(`Crew Presence Board API läuft auf Port ${PORT}`));

teamsNotifications.start();
