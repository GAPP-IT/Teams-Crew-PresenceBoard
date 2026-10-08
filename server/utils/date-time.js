export function formatTime(value) {
  if (!value) return null;
  const d = new Date(value);
  return Number.isNaN(d.getTime())
    ? null
    : new Intl.DateTimeFormat("de-DE", { hour: "2-digit", minute: "2-digit", hour12: false, timeZone: "Europe/Berlin" }).format(d);
}
export function formatDate(value) {
  if (!value) return null;
  const d = new Date(value);
  return Number.isNaN(d.getTime())
    ? null
    : new Intl.DateTimeFormat("de-DE", { day: "2-digit", month: "2-digit", year: "numeric", timeZone: "Europe/Berlin" }).format(d);
}
export function formatDateTime(value) {
  if (!value) return null;
  const d = new Date(value);
  return Number.isNaN(d.getTime())
    ? null
    : new Intl.DateTimeFormat("de-DE", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        timeZone: "Europe/Berlin",
      }).format(d);
}
export function parseGraphDateTime(value) {
  const dateTime = typeof value === "string" ? value : value?.dateTime;
  if (!dateTime) return new Date(Number.NaN);
  if (/(?:Z|[+-]\d{2}:\d{2})$/i.test(dateTime)) return new Date(dateTime);
  const wallTime = new Date(`${dateTime}Z`);
  if (Number.isNaN(wallTime.getTime())) return wallTime;
  const rawZone = typeof value === "object" ? value.timeZone || "UTC" : "UTC";
  const zoneAliases = { "w. europe standard time": "Europe/Berlin", "romance standard time": "Europe/Paris" };
  const timeZone = zoneAliases[rawZone.toLowerCase()] || rawZone;
  if (["utc", "gmt", "etc/utc"].includes(timeZone.toLowerCase())) return wallTime;
  try {
    const formatter = new Intl.DateTimeFormat("en-US", {
      timeZone,
      hourCycle: "h23",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
    });
    let timestamp = wallTime.getTime();
    for (let attempt = 0; attempt < 2; attempt++) {
      const parts = Object.fromEntries(formatter.formatToParts(new Date(timestamp)).map((part) => [part.type, part.value]));
      const representedAsUtc = Date.UTC(
        Number(parts.year),
        Number(parts.month) - 1,
        Number(parts.day),
        Number(parts.hour),
        Number(parts.minute),
        Number(parts.second),
      );
      timestamp = wallTime.getTime() - (representedAsUtc - timestamp);
    }
    return new Date(timestamp);
  } catch {
    return new Date(dateTime);
  }
}
