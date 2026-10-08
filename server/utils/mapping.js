import {
  formatDate,
  formatDateTime,
  formatTime,
  parseGraphDateTime,
} from "./date-time.js";

export function mapCalendar(events) {
  const now = new Date();
  const list = (events || [])
    .filter((e) => !e.isCancelled)
    .map((e) => ({ ...e, s: parseGraphDateTime(e.start), e: parseGraphDateTime(e.end) }))
    .filter((e) => !Number.isNaN(e.s.getTime()) && !Number.isNaN(e.e.getTime()))
    .sort((a, b) => a.s - b.s);
  const today = formatDate(now);
  const oof = list.find((e) => e.showAs === "oof" && e.s <= now && e.e > now);
  const current = list.find((e) => e.showAs !== "oof" && !e.isAllDay && e.s <= now && e.e > now);
  const next = list.find((e) => e.showAs !== "oof" && !e.isAllDay && e.s > now && formatDate(e.s) === today);
  const schedule = list.map((event) => ({
    start: event.s.toISOString(),
    end: event.e.toISOString(),
    showAs: event.showAs || "busy",
    isAllDay: Boolean(event.isAllDay),
  }));
  return {
    currentMeetingEnd: current ? formatTime(current.e) : null,
    currentMeetingShowAs: current?.showAs || null,
    meetingReachable: current ? current.showAs === "free" : true,
    nextMeeting: !current && next ? formatTime(next.s) : null,
    oofAbsence: oof ? { returnDate: formatDate(oof.e) } : null,
    schedule,
  };
}

export function mapAutomaticReplies(settings) {
  if (!settings) return { configured: false, status: "unavailable", active: false, note: "" };
  const replies = settings.automaticRepliesSetting || settings;
  if (!replies) return { configured: false, status: "missing", active: false, note: "" };
  const status = (replies?.status || "").trim().toLowerCase();
  const startDateTime = replies.scheduledStartDateTime;
  const endDateTime = replies.scheduledEndDateTime;
  const internalNote = plainText(replies.internalReplyMessage || "");
  const externalNote = plainText(replies.externalReplyMessage || replies.externalRelayMessage || "");
  const note = internalNote || externalNote;
  if (status === "disabled") return { configured: true, status, active: false, note };
  if (status === "alwaysenabled" || status === "enabled") return { configured: true, status, active: true, returnDate: null, note };
  if (status !== "scheduled") return { configured: false, status: status || "unknown", active: false, note };
  const hasStart = Boolean(startDateTime?.dateTime);
  const hasEnd = Boolean(endDateTime?.dateTime);
  const start = parseGraphDateTime(startDateTime);
  const end = parseGraphDateTime(endDateTime);
  const now = new Date();
  if (!hasStart && !hasEnd) return { configured: true, status, active: true, returnDate: null, note };
  if ((hasStart && Number.isNaN(start.getTime())) || (hasEnd && Number.isNaN(end.getTime())))
    return { configured: true, status, active: false, note };
  if ((hasStart && start > now) || (hasEnd && end <= now)) return { configured: true, status, active: false, note };
  return { configured: true, status, active: true, returnDate: hasEnd ? formatDateTime(end) : null, note };
}
export function mapPresence(p) {
  const availability = p?.availability || "PresenceUnknown",
    activity = p?.activity || "Offline";
  const pm = {
    Available: "available",
    AvailableIdle: "available",
    Busy: "busy",
    BusyIdle: "busy",
    DoNotDisturb: "dnd",
    Away: "away",
    BeRightBack: "away",
    Offline: "offline",
    PresenceUnknown: "offline",
  };
  const outOfOffice = [availability, activity].some((value) => value.toLowerCase().replace(/[\s_-]/g, "") === "outofoffice");
  let presence = pm[availability] || "offline";
  if (!outOfOffice && ["InAMeeting", "Presenting"].includes(activity)) presence = "meeting";
  const phone = ["InACall", "InAConferenceCall"].includes(activity) ? "call" : "free";
  const rawLocationType = p?.workLocation?.workLocationType || p?.workLocation?.type || p?.workLocationType || p?.location || "";
  const type = typeof rawLocationType === "string" ? rawLocationType.toLowerCase().replace(/[\s_-]/g, "") : "";
  const location = ["remote", "home", "homeoffice", "workfromhome", "wfh"].includes(type)
    ? "home"
    : ["office", "onsite", "onpremises"].includes(type)
      ? "office"
      : null;
  return { presence, phone, availability, activity, location, outOfOffice };
}

export function plainText(value = "") {
  return String(value)
    .replace(/<br\s*\/?\s*>/gi, " ")
    .replace(/<\/p\s*>/gi, " ")
    .replace(/<[^>]*>/g, " ")
    .replace(/&nbsp;|&#160;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/\s+/g, " ")
    .trim();
}