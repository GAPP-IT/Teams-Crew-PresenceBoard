import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import * as teamsJs from "@microsoft/teams-js";
import {
  ArrowLeft, Briefcase, Building2, CalendarDays, ChevronDown, ChevronsDown, ChevronsUp, Folder, Headphones, Heart,
  Home, Info, Languages, Mail, MapPin, MapPinOff, MessageCircle, Moon, Pencil, Phone, PhoneOff,
  RefreshCw, Search, SlidersHorizontal, Plus, Star, Sun, Tag, Trash2, Users, Video, VideoOff, X, Bell, BellOff
} from "lucide-react";
import { translate } from "./translations.js";
import { apiFetch } from "./services/api";

const API = import.meta.env.VITE_API_BASE_URL || "/api";
const USER_GUIDE_URL = "";
const FAVORITES_KEY = "crewPresenceBoardFavorites";
const FAVORITE_GROUPS_KEY = "crewPresenceBoardGroups";
const FAVORITE_GROUP_COLLAPSED_KEY = "crewPresenceBoardCollapsedFavoriteGroups";
const THEME_KEY = "crewPresenceBoardTheme";
const VIEW_KEY = "crewPresenceBoardDefaultView";
const LANGUAGE_KEY = "crewPresenceBoardLanguage";
const PRESENCE_NOTIFICATIONS_KEY = "crewPresenceBoardPresenceNotifications";
const MEETING_ROOM_COMPANY = "__meeting_rooms__";
const MISSING_COMPANY = "__missing_company__";
const MISSING_DEPARTMENT = "__missing_department__";
const MISSING_ROLE = "__missing_role__";

const companies = [
  { id: "PAG", name: "Piper Deutschland AG", color: "#9A2F38" },
  { id: "GAPP", name: "Global Aviation + Piper Parts GmbH", color: "#486893" },
  { id: "EAC2", name: "European Aviation Competence Center GmbH", color: "#F1880F" },
  { id: MEETING_ROOM_COMPANY, name: "Besprechungsräume", color: "#83d0e7" },
  { id: MISSING_COMPANY, name: "Ohne Firma", color: "#727C8A" }
];

const presenceMap = {
  available: { labelKey: "presenceAvailable", color: "#35A853" },
  busy: { labelKey: "presenceBusy", color: "#C4314B" },
  meeting: { labelKey: "presenceMeeting", color: "#8B5CF6" },
  dnd: { labelKey: "presenceDnd", color: "#B00020" },
  away: { labelKey: "presenceAway", color: "#F59E0B" },
  offline: { labelKey: "presenceOffline", color: "#8A8A8A" }
};

const phoneMap = {
  free: { labelKey: "phoneFree", color: "#35A853", icon: Phone },
  call: { labelKey: "phoneCall", color: "#D83B01", icon: Headphones }
};

const favoriteGroupIcons = {
  folder: Folder,
  star: Star,
  users: Users,
  briefcase: Briefcase,
  heart: Heart,
  tag: Tag,
  building: Building2
};

const favoriteGroupIconOptions = [
  { id: "folder", labelKey: "favoriteGroupIconFolder" },
  { id: "star", labelKey: "favoriteGroupIconStar" },
  { id: "users", labelKey: "favoriteGroupIconUsers" },
  { id: "briefcase", labelKey: "favoriteGroupIconBriefcase" },
  { id: "heart", labelKey: "favoriteGroupIconHeart" },
  { id: "tag", labelKey: "favoriteGroupIconTag" },
  { id: "building", labelKey: "favoriteGroupIconBuilding" }
];

function FavoriteGroupIcon({ name, color }) {
  const Icon = favoriteGroupIcons[name] || Folder;
  return <Icon size={14} style={{ color: color || "#64748B" }} />;
}

function readStored(key, fallback) {
  try {
    const value = localStorage.getItem(key);
    return value === null ? fallback : JSON.parse(value);
  } catch {
    return fallback;
  }
}
function store(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {}
}

function companyId(value = "") {
  const s = String(value || "").trim().toLowerCase();
  if (s.includes("meeting room") || s === MEETING_ROOM_COMPANY) return MEETING_ROOM_COMPANY;
  if (!s || s === MISSING_COMPANY) return MISSING_COMPANY;
  if (s === "gapp" || s.includes("global aviation")) return "GAPP";
  if (s === "eac2" || s === "eacc" || s.includes("european aviation")) return "EAC2";
  return "PAG";
}
function initials(person) {
  if (person.initials) return person.initials;
  const n = person.name || "";
  if (n.includes(",")) {
    const [last, first] = n.split(",").map((x) => x.trim());
    return `${first?.[0] || ""}${last?.[0] || ""}`.toUpperCase();
  }
  return n
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((x) => x[0])
    .join("")
    .toUpperCase();
}
function sortPeople(people, sortOrder, language) {
  const nameOrder = new Intl.Collator(language, { numeric: true, sensitivity: "base" });
  const availabilityOrder = {
    available: 0,
    away: 1,
    busy: 2,
    dnd: 2,
    meeting: 3,
    offline: 4
  };

  return [...people].sort((left, right) => {
    if (sortOrder === "name-desc") {
      return nameOrder.compare(right.name, left.name);
    }
    if (sortOrder === "availability") {
      const leftRank = left.sageAbsence || left.oofAbsence
        ? availabilityOrder.offline
        : availabilityOrder[left.presence] ?? availabilityOrder.offline;
      const rightRank = right.sageAbsence || right.oofAbsence
        ? availabilityOrder.offline
        : availabilityOrder[right.presence] ?? availabilityOrder.offline;
      if (leftRank !== rightRank) return leftRank - rightRank;
    }
    return nameOrder.compare(left.name, right.name);
  });
}
function normalize(person) {
  const activity = (person.activity || "").toLowerCase();
  const rawLocation = (person.location || "").toLowerCase();
  const locationType = rawLocation.replace(/[\s_-]/g, "");
  return {
    ...person,
    name: person.name || person.displayName || "",
    initials: initials(person),
    company: companyId(person.companyName || person.company),
    department: person.department || MISSING_DEPARTMENT,
    role: person.role || MISSING_ROLE,
    presence: person.presence || "offline",
    phone:
      person.phone === "call" ||
      activity === "inacall" ||
      activity === "inaconferencecall"
        ? "call"
        : "free",
    location: ["home", "remote", "homeoffice", "workfromhome", "wfh"].includes(locationType)
      ? "home"
      : ["office", "onsite", "onpremises"].includes(locationType)
        ? "office"
        : null,
    photoUrl: person.photoUrl || `${API}/users/${person.id}/photo`
  };
}
function notificationPhoneStatus(person) {
  if (person.presence === "away") return "away";
  if (person.phone === "call") return "call";
  if (["available", "busy", "meeting"].includes(person.presence)) return "free";
  return "other";
}
function deepLink(kind, email = "") {
  const user = encodeURIComponent(email);
  if (kind === "chat") return `https://teams.microsoft.com/l/chat/0/0?users=${user}`;
  if (kind === "call") return `https://teams.microsoft.com/l/call/0/0?users=${user}`;
  if (kind === "video") return `https://teams.microsoft.com/l/call/0/0?users=${user}&withVideo=true`;
  return `mailto:${email}`;
}

function Logo() {
  return (
    <img
      className="app-logo"
      src="/PiperLogo_Black.png"
      alt="Piper Deutschland AG"
    />
  );
}

function IconButton({ icon: Icon, title, ariaLabel, disabled, onClick, t, className = "" }) {
  return (
    <button
      className={`icon-button ${className}`}
      type="button"
      title={disabled ? t("actionUnavailable", { action: title }) : title}
      aria-label={ariaLabel || title}
      disabled={disabled}
      onClick={onClick}
    >
      <Icon size={15} />
    </button>
  );
}

function Photo({ person, company }) {
  const [failed, setFailed] = useState(false);
  if (failed) {
    return (
      <div
        className="avatar fallback"
        style={{ background: `linear-gradient(145deg,${company.color},${company.color}bb)` }}
      >
        {person.initials}
      </div>
    );
  }
  return (
    <img
      className="avatar"
      src={person.photoUrl}
      alt=""
      onError={() => setFailed(true)}
    />
  );
}

function PersonCard({
  person,
  company,
  favorite,
  onFavorite,
  watchingPresence,
  onWatchPresence,
  t
}) {
  const absence = person.sageAbsence || person.oofAbsence;
  const returnDate = person.oofAbsence?.returnDate || person.sageAbsence?.returnDate;
  const absenceNote = person.oofAbsence?.note || person.outOfOfficeNote;
  const absenceTitle =
    [
      absenceNote || t("noAbsenceNote"),
      returnDate ? t("returnOn", { date: returnDate }) : t("returnDateUnset")
    ].join("\n");
  const presenceStatus = presenceMap[person.presence] || presenceMap.offline;
  const p = {
    ...(absence ? presenceMap.offline : presenceStatus),
    label: t((absence ? presenceMap.offline : presenceStatus).labelKey)
  };
  const phoneBarColor = p.color;
  const isOnCall = person.phone === "call" || person.presence === "meeting";
  const teamsAway = !absence && person.presence === "away";
  const phoneStatus = person.phone === "call"
    ? phoneMap.call
    : teamsAway
    ? { labelKey: "phoneFreeTeamsAway", color: "#F59E0B", icon: Phone }
    : person.presence === "dnd"
      ? { labelKey: "phoneUnavailable", color: "#C4314B", icon: PhoneOff }
      : phoneMap[person.phone] || phoneMap.free;
  const PhoneIcon = phoneStatus.icon;
  const phoneLabel = t(phoneStatus.labelKey);
  const offline = Boolean(absence) || person.presence === "offline";
  const callDisabled = Boolean(absence) || offline || person.presence === "dnd" || isOnCall;
  const showPhone = !absence && !offline;
  const showPhoneStatus = showPhone && person.presence !== "meeting";
  const location =
    offline
      ? null
      : person.location === "home"
        ? { label: t("homeOffice"), icon: Home }
        : person.location === "office"
          ? { label: t("office"), icon: Building2 }
          : { label: t("locationUnavailable"), icon: MapPinOff };
  const LocationIcon = location?.icon;

  return (
    <article className="person-card">
      <div className="phone-bar" style={{ backgroundColor: phoneBarColor }} />
      <div className="avatar-wrap">
        <Photo person={person} company={company} />
        <span
          className="presence-dot"
          style={{ backgroundColor: p.color }}
          title={p.label}
        />
      </div>
      <div className="person-content">
        <div className="name-row">
          <strong title={person.name}>{person.name}</strong>
          <button
            className="favorite-button"
            onClick={() => onFavorite(person.id)}
            title={favorite ? t("favoriteRemove") : t("favoriteAdd")}
          >
            <Star
              size={13}
              className={favorite ? "favorite-active" : "favorite-inactive"}
            />
          </button>
        </div>
        <div className="role" title={person.role}>
          {person.role === MISSING_ROLE ? t("roleMissing") : person.role}
        </div>
        <div className="status-row">
          {absence ? (
            <span
              className="detail absence"
              title={absenceTitle}
              aria-label={absenceTitle}
            >
              <ArrowLeft size={11} />
              {returnDate
                ? t("returnOn", { date: returnDate })
                : t("returnDateUnset")}
            </span>
          ) : (
            <>
              <span className="status-item" style={{ color: p.color }}>
                <i style={{ backgroundColor: p.color }} />
                {p.label}
              </span>
              {showPhoneStatus && (
                <span className="status-item" style={{ color: phoneStatus.color }}>
                  <PhoneIcon size={11} />
                  {phoneLabel}
                </span>
              )}
            </>
          )}
        </div>
        {(location ||
          (!absence && (person.currentMeetingEnd || person.nextMeeting))) && (
          <div className="details-row">
            {location && (
              <span className="detail" title={location.label} aria-label={location.label}>
                <LocationIcon size={14} style={{ flexShrink: 0 }} />
              </span>
            )}
            {!absence && person.currentMeetingEnd && (
              <span className="detail meeting">
                <CalendarDays className="appointment-icon" size={11} />
                {t("appointmentEntry", { time: person.currentMeetingEnd })}
              </span>
            )}
            {!absence && !person.currentMeetingEnd && person.nextMeeting && (
              <span className="detail next-appointment">
                <CalendarDays className="appointment-icon" size={11} />
                {t("nextAppointment", { time: person.nextMeeting })}
              </span>
            )}
          </div>
        )}
      </div>
      <div className="card-actions">
        {/*
        <IconButton
          icon={watchingPresence ? BellOff : Bell}
          title={watchingPresence ? t("presenceNotificationStop") : t("presenceNotificationStart")}
          onClick={() => onWatchPresence(person)}
          className={watchingPresence ? "is-watching" : ""}
          t={t}
        />
        */}
        <IconButton
          icon={MessageCircle}
          title={t("chatStart")}
          onClick={() => window.open(deepLink("chat", person.email), "_blank", "noopener,noreferrer")}
          t={t}
        />
        <IconButton
          icon={callDisabled ? PhoneOff : Phone}
          title={t("call")}
          disabled={callDisabled}
          onClick={() => window.open(deepLink("call", person.email), "_blank", "noopener,noreferrer")}
          t={t}
        />
        <IconButton
          icon={callDisabled ? VideoOff : Video}
          title={t("videoCall")}
          disabled={callDisabled}
          onClick={() => window.open(deepLink("video", person.email), "_blank", "noopener,noreferrer")}
          t={t}
        />
        <IconButton
          icon={Mail}
          title={t("email")}
          onClick={() => {
            window.location.href = deepLink("mail", person.email);
          }}
          t={t}
        />
      </div>
    </article>
  );
}

function formatResourceDateTime(value, language) {
  return new Intl.DateTimeFormat(language === "de" ? "de-DE" : "en-GB", {
    weekday: "short",
    day: "2-digit",
    month: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  }).format(new Date(value));
}

function ResourceCard({ person, onOpenCalendar, onOpenInfo, t, language }) {
  const schedule = Array.isArray(person.resourceSchedule) ? person.resourceSchedule : [];
  const calendarAvailable = person.resourceCalendarAvailable !== false;
  const occupied = schedule.filter((event) => event.showAs !== "free");
  const [now, setNow] = useState(() => Date.now());

  useEffect(() => {
    let timeoutId;
    const updateAtNextBookingBoundary = () => {
      const currentTime = Date.now();
      setNow(currentTime);
      const nextBoundary = schedule
        .flatMap((event) => [new Date(event.start).getTime(), new Date(event.end).getTime()])
        .filter((timestamp) => Number.isFinite(timestamp) && timestamp > currentTime)
        .sort((left, right) => left - right)[0];

      if (nextBoundary) {
        timeoutId = window.setTimeout(updateAtNextBookingBoundary, nextBoundary - currentTime + 10);
      }
    };

    updateAtNextBookingBoundary();
    return () => window.clearTimeout(timeoutId);
  }, [schedule]);

  const currentBooking = calendarAvailable && occupied.find(
    (event) => new Date(event.start).getTime() <= now && new Date(event.end).getTime() > now
  );
  const nextBooking = calendarAvailable && occupied.find((event) => new Date(event.start).getTime() > now);
  const statusKey = !calendarAvailable
    ? "resourceStatusUnavailable"
    : currentBooking
      ? "resourceOccupied"
      : "resourceAvailable";
  const statusColor = !calendarAvailable
    ? "#747D8A"
    : currentBooking
      ? "#C4314B"
      : "#35A853";
  const detailText = !calendarAvailable
    ? t("resourceCalendarUnavailable")
    : currentBooking
      ? t("resourceOccupiedUntil", {
          time: formatResourceDateTime(currentBooking.end, language)
        })
      : nextBooking
        ? t("resourceNextBooking", {
            time: formatResourceDateTime(nextBooking.start, language)
          })
        : t("resourceNoUpcoming");

  return (
    <article className="person-card resource-person-card">
      <div className="phone-bar" style={{ backgroundColor: statusColor }} />
      <div className="avatar-wrap resource-avatar-wrap">
        <div className="avatar resource-avatar" aria-hidden="true">
          <Building2 size={19} />
        </div>
      </div>
      <div className="person-content">
        <div className="name-row">
          <strong title={person.name}>{person.name}</strong>
        </div>
        <div className="role">{t("meetingRoomCompany")}</div>
        <div className="status-row resource-status-row">
          <span className="status-item" style={{ color: statusColor }}>
            <i style={{ backgroundColor: statusColor }} />
            {t(statusKey)}
          </span>
          {person.resourceCapacity != null && (
            <span className="status-item resource-capacity" title={t("resourceCapacity", { count: person.resourceCapacity })}>
              <Users size={13} />
              {t("resourceCapacity", { count: person.resourceCapacity })}
            </span>
          )}
        </div>
        {person.resourceLocationDescription && (
          <div className="details-row resource-meta-row">
            <span className="detail" title={person.resourceLocationDescription}>
              <MapPin size={13} />
              <span>{person.resourceLocationDescription}</span>
            </span>
          </div>
        )}
        <div className="details-row">
          <span className="detail meeting" title={detailText}>
            <CalendarDays className="appointment-icon" size={11} />
            {detailText}
          </span>
        </div>
      </div>
      <div className="card-actions">
        <IconButton
          icon={Info}
          title={t("resourceOpenInfo")}
          ariaLabel={`${t("resourceOpenInfo")}: ${person.name}`}
          onClick={() => onOpenInfo(person)}
          t={t}
        />
        <IconButton
          icon={CalendarDays}
          title={t("resourceOpenCalendar")}
          ariaLabel={`${t("resourceOpenCalendar")}: ${person.name}`}
          onClick={() => onOpenCalendar(person)}
          t={t}
        />
      </div>
    </article>
  );
}

function ResourceCalendarModal({ person, onClose, t, language }) {
  if (!person) return null;
  const schedule = Array.isArray(person.resourceSchedule) ? person.resourceSchedule : [];
  const upcomingBookings = schedule
    .filter((event) => event.showAs !== "free" && new Date(event.end).getTime() > Date.now())
    .sort((left, right) => new Date(left.start).getTime() - new Date(right.start).getTime())
    .slice(0, 10);

  return (
    <div className="dialog-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}>
      <div className="info-dialog resource-calendar-dialog" role="dialog" aria-modal="true" aria-label={person.name}>
        <button className="dialog-close" type="button" onClick={onClose} aria-label={t("close")}>
          <X size={18} />
        </button>
        <h2>{person.name}</h2>
        <p className="resource-calendar-subtitle">{t("resourceCalendarTitle")}</p>
        {person.resourceCalendarAvailable === false ? (
          <p className="resource-calendar-empty">{t("resourceCalendarUnavailable")}</p>
        ) : upcomingBookings.length ? (
          <div className="resource-calendar-list">
            {upcomingBookings.map((event, index) => (
              <div className="resource-calendar-entry" key={`${event.start}-${index}`}>
                <span className={`resource-calendar-dot ${event.showAs === "free" ? "is-free" : "is-busy"}`} />
                <div>
                  <strong>
                    {event.isAllDay
                      ? t("resourceAllDay")
                      : `${formatResourceDateTime(event.start, language)} – ${formatResourceDateTime(event.end, language)}`}
                  </strong>
                  <span>{t(event.showAs === "free" ? "resourceAvailable" : "resourceOccupied")}</span>
                </div>
              </div>
            ))}
          </div>
        ) : (
          <p className="resource-calendar-empty">{t("resourceCalendarEmpty")}</p>
        )}
        <footer>
          <button className="outline-button" type="button" onClick={onClose}>{t("close")}</button>
        </footer>
      </div>
    </div>
  );
}

function ResourceInfoModal({ person, onClose, t }) {
  const [photoUnavailable, setPhotoUnavailable] = useState(false);
  if (!person) return null;
  const profilePhotoUrl = person.photoUrl || `${API}/users/${person.id}/photo`;

  return (
    <div className="dialog-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}>
      <div className="resource-calendar-dialog resource-info-dialog" role="dialog" aria-modal="true" aria-label={person.name}>
        <button className="dialog-close" type="button" onClick={onClose} aria-label={t("close")}>
          <X size={18} />
        </button>
        <h2>{person.name}</h2>
        <p className="resource-calendar-subtitle">{t("resourceInfoTitle")}</p>
        {photoUnavailable ? (
          <div className="resource-info-photo-placeholder" aria-label={t("resourcePhotoUnavailable")}>
            <Building2 size={38} />
          </div>
        ) : (
          <img
            className="resource-info-photo"
            src={profilePhotoUrl}
            alt={`${t("meetingRoomCompany")}: ${person.name}`}
            onError={() => {
              setPhotoUnavailable(true);
            }}
          />
        )}
        <dl className="resource-info-details">
          {person.resourceLocationDescription && (
            <div>
              <dt>{t("resourceLocation")}</dt>
              <dd>{person.resourceLocationDescription}</dd>
            </div>
          )}
          {person.resourceCapacity != null && (
            <div>
              <dt>{t("resourceCapacityLabel")}</dt>
              <dd>{t("resourceCapacity", { count: person.resourceCapacity })}</dd>
            </div>
          )}
        </dl>
        <footer>
          <button className="outline-button" type="button" onClick={onClose}>{t("close")}</button>
        </footer>
      </div>
    </div>
  );
}

function InfoModal({ open, onClose, defaultView, onDefaultView, t }) {
  if (!open) return null;
  return (
    <div
      className="modal-backdrop"
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="info-modal" role="dialog" aria-modal="true">
        <header>
          <div>
            <h2>Crew Presence Board</h2>
            <p>{t("information")}</p>
          </div>
          <button className="header-icon-button" onClick={onClose}>
            <X size={18} />
          </button>
        </header>
        <div className="modal-content">
          <h3>{t("aboutTitle")}</h3>
          <p>
            {t("aboutDescription")}
          </p>
          <p className="user-guide-link-row">
            <a
              className={`user-guide-link${USER_GUIDE_URL ? "" : " is-placeholder"}`}
              href={USER_GUIDE_URL || undefined}
              target={USER_GUIDE_URL ? "_blank" : undefined}
              rel="noreferrer"
              role="link"
              tabIndex={USER_GUIDE_URL ? undefined : 0}
              aria-disabled={!USER_GUIDE_URL}
              onClick={(event) => {
                if (!USER_GUIDE_URL) event.preventDefault();
              }}
            >
              {t("userGuide")}
              {!USER_GUIDE_URL && <span>{t("userGuideLinkPending")}</span>}
            </a>
          </p>
          <h3>{t("updatesTitle")}</h3>
          <p>
            {t("updatesDescription")}
          </p>
        </div>
        <footer>
          <button className="outline-button" onClick={onClose}>
            {t("close")}
          </button>
        </footer>
      </div>
    </div>
  );
}

export default function App() {
  const [people, setPeople] = useState([]);
  const [favorites, setFavorites] = useState(() => readStored(FAVORITES_KEY, []));
  const [presenceNotifications, setPresenceNotifications] = useState(() => {
    const stored = readStored(PRESENCE_NOTIFICATIONS_KEY, []);
    return Array.isArray(stored) ? stored.map(String) : [];
  });
  const [teamsContext, setTeamsContext] = useState({ ready: false, host: false });
  const [teamsSubscriptionsReady, setTeamsSubscriptionsReady] = useState(false);
  const [notificationToast, setNotificationToast] = useState(null);
  const [favoriteGroupsData, setFavoriteGroupsData] = useState(() => {
    const stored = readStored(FAVORITE_GROUPS_KEY, {});
    return {
      groups: Array.isArray(stored?.groups) ? stored.groups : [],
      assignments:
        stored?.assignments && typeof stored.assignments === "object"
          ? stored.assignments
          : {}
    };
  });
  const [collapsedFavoriteGroups, setCollapsedFavoriteGroups] = useState(() => {
    const stored = readStored(FAVORITE_GROUP_COLLAPSED_KEY, {});
    return stored && typeof stored === "object" && !Array.isArray(stored) ? stored : {};
  });
  const [collapsedCompanies, setCollapsedCompanies] = useState({});
  const [collapsedDepartments, setCollapsedDepartments] = useState({});
  const [newFavoriteGroupName, setNewFavoriteGroupName] = useState("");
  const [editingFavoriteGroupId, setEditingFavoriteGroupId] = useState(null);
  const [favoriteGroupDraft, setFavoriteGroupDraft] = useState(null);
  const [addingToFavoriteGroup, setAddingToFavoriteGroup] = useState(null);
  const [favoriteToAdd, setFavoriteToAdd] = useState("");
  const [theme, setTheme] = useState(() => readStored(THEME_KEY, "light"));
  const [language, setLanguage] = useState(() =>
    readStored(LANGUAGE_KEY, "de") === "en" ? "en" : "de"
  );
  const [sortOrder, setSortOrder] = useState("availability");
  const [defaultView, setDefaultView] = useState(() => readStored(VIEW_KEY, "all"));
  const [view, setView] = useState(() => readStored(VIEW_KEY, "all"));
  const [query, setQuery] = useState("");
  const [department, setDepartment] = useState("all");
  const [presence, setPresence] = useState("all");
  const [phone, setPhone] = useState("all");
  const [location, setLocation] = useState("all");
  const [filtersOpen, setFiltersOpen] = useState(false);
  const [infoOpen, setInfoOpen] = useState(false);
  const [resourceCalendarPerson, setResourceCalendarPerson] = useState(null);
  const [resourceInfoPerson, setResourceInfoPerson] = useState(null);
  const [loading, setLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const [error, setError] = useState(null);
  const [lastUpdated, setLastUpdated] = useState(null);
  const previousPeopleRef = useRef(null);
  const notificationRegistrationRef = useRef(null);
  const notificationToastTimeoutRef = useRef(null);
  const teamsSubscriptionQueueRef = useRef(Promise.resolve());
  const teamsNotificationErrorReportedRef = useRef(false);
  const loadInProgressRef = useRef(false);
  const t = useMemo(
    () => (key, values) => translate(language, key, values),
    [language]
  );
  const syncTeamsSubscriptions = useCallback(async (method, personIds) => {
    const accessToken = await teamsJs.authentication.getAuthToken();
    const response = await fetch(`${API}/teams/notifications/subscriptions`, {
      method,
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
        ...(method === "PUT" ? { "Content-Type": "application/json" } : {})
      },
      ...(method === "PUT" ? { body: JSON.stringify({ personIds }) } : {})
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Teams notification setup failed.");
    teamsNotificationErrorReportedRef.current = false;
    return data;
  }, []);
  const reportTeamsNotificationError = useCallback((error) => {
    console.warn("Teams notifications could not be synchronized.", error);
    if (teamsNotificationErrorReportedRef.current) return;
    teamsNotificationErrorReportedRef.current = true;
    setNotificationToast({
      title: t("teamsNotificationSetupTitle"),
      body: t("teamsNotificationSetupBody")
    });
  }, [t]);
  const enqueueTeamsSubscriptionSync = useCallback((operation) => {
    teamsSubscriptionQueueRef.current = teamsSubscriptionQueueRef.current
      .catch(() => {})
      .then(operation);
    return teamsSubscriptionQueueRef.current;
  }, []);

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    store(THEME_KEY, theme);
  }, [theme]);
  useEffect(() => {
    document.documentElement.lang = language;
    store(LANGUAGE_KEY, language);
  }, [language]);
  useEffect(() => {
    if (!("serviceWorker" in navigator)) return undefined;
    navigator.serviceWorker
      .register("/notification-sw.js")
      .then((registration) => {
        notificationRegistrationRef.current = registration;
      })
      .catch((registrationError) => {
        console.warn("Service Worker für Benachrichtigungen konnte nicht registriert werden.", registrationError);
      });
    return undefined;
  }, []);
  useEffect(() => {
    let active = true;
    teamsJs.app.initialize()
      .then(() => teamsJs.app.getContext())
      .then((context) => {
        if (active) {
          setTeamsContext({ ready: true, host: context.app?.host?.name === "teams" });
        }
      })
      .catch(() => {
        if (active) setTeamsContext({ ready: true, host: false });
      });
    return () => {
      active = false;
    };
  }, []);
  useEffect(() => {
    if (!teamsContext.ready || !teamsContext.host) return undefined;
    let active = true;
    enqueueTeamsSubscriptionSync(async () => {
      const data = await syncTeamsSubscriptions("GET");
      if (!active) return;
      setPresenceNotifications((current) =>
        [...new Set([...data.personIds, ...current.map(String)])]
      );
      setTeamsSubscriptionsReady(true);
    }).catch((error) => {
      if (active) reportTeamsNotificationError(error);
    });
    return () => {
      active = false;
    };
  }, [teamsContext.ready, teamsContext.host, syncTeamsSubscriptions, enqueueTeamsSubscriptionSync, reportTeamsNotificationError]);
  useEffect(() => {
    if (!teamsContext.host || !teamsSubscriptionsReady) return undefined;
    let active = true;
    const personIds = presenceNotifications.map(String);
    enqueueTeamsSubscriptionSync(() => syncTeamsSubscriptions("PUT", personIds))
      .catch((error) => {
        if (active) reportTeamsNotificationError(error);
      });
    return () => {
      active = false;
    };
  }, [teamsContext.host, teamsSubscriptionsReady, presenceNotifications, syncTeamsSubscriptions, enqueueTeamsSubscriptionSync, reportTeamsNotificationError]);
  useEffect(() => {
    if (!teamsContext.host || !teamsSubscriptionsReady || !lastUpdated) return undefined;
    let active = true;
    enqueueTeamsSubscriptionSync(async () => {
      const data = await syncTeamsSubscriptions("GET");
      if (!active) return;
      setPresenceNotifications((current) => {
        const currentIds = current.map(String);
        return currentIds.length === data.personIds.length &&
          currentIds.every((id) => data.personIds.includes(id))
          ? currentIds
          : data.personIds;
      });
    }).catch((error) => {
      if (active) reportTeamsNotificationError(error);
    });
    return () => {
      active = false;
    };
  }, [teamsContext.host, teamsSubscriptionsReady, lastUpdated, syncTeamsSubscriptions, enqueueTeamsSubscriptionSync, reportTeamsNotificationError]);
  useEffect(() => store(FAVORITES_KEY, favorites), [favorites]);
  useEffect(
    () => store(PRESENCE_NOTIFICATIONS_KEY, presenceNotifications),
    [presenceNotifications]
  );
  useEffect(() => store(FAVORITE_GROUPS_KEY, favoriteGroupsData), [favoriteGroupsData]);
  useEffect(
    () => store(FAVORITE_GROUP_COLLAPSED_KEY, collapsedFavoriteGroups),
    [collapsedFavoriteGroups]
  );
  useEffect(() => store(VIEW_KEY, defaultView), [defaultView]);

  const showPresenceNotification = useCallback(async (title, body, personId) => {
    if (typeof Notification !== "undefined" && Notification.permission === "granted") {
      try {
        const registration = notificationRegistrationRef.current;
        if (registration?.showNotification) {
          await registration.showNotification(title, {
            body,
            icon: "/PiperLogo_Black.png",
            tag: `presence-${personId}`
          });
          return;
        }
        new Notification(title, { body, icon: "/PiperLogo_Black.png" });
        return;
      } catch (notificationError) {
        console.warn("Systembenachrichtigung fehlgeschlagen; zeige Hinweis in der App.", notificationError);
      }
    }
    setNotificationToast({ title, body });
    window.clearTimeout(notificationToastTimeoutRef.current);
    notificationToastTimeoutRef.current = window.setTimeout(() => {
      setNotificationToast(null);
    }, 7000);
  }, []);

  const load = useCallback(async (showSpinner = false) => {
    if (loadInProgressRef.current) return;
    loadInProgressRef.current = true;
    try {
      if (showSpinner) setRefreshing(true);
      /*const response = await fetch(`${API}/presence-board`, {
        cache: "no-store",
        headers: { Accept: "application/json" }
      });*/

        const response =
          await apiFetch("/api/presence-board");

        if (!response.ok) {
          const error = await response
            .json()
            .catch(() => null);

          throw new Error(
            error?.error ||
              `HTTP ${response.status}`,
          );
        }



      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.details || data.error || "API konnte nicht geladen werden.");
      }
      const nextPeople = (data.people || []).map(normalize);
      const previousPeople = previousPeopleRef.current;
      if (previousPeople && !teamsContext.host) {
        const previousById = new Map(previousPeople.map((person) => [String(person.id), person]));
        const changedWatchedPeople = nextPeople.filter((person) => {
          const previous = previousById.get(String(person.id));
          const previousPhoneStatus = previous && notificationPhoneStatus(previous);
          const nextPhoneStatus = notificationPhoneStatus(person);
          return (
            presenceNotifications.includes(String(person.id)) &&
            previous &&
            ["call", "away"].includes(previousPhoneStatus) &&
            nextPhoneStatus === "free"
          );
        });
        if (changedWatchedPeople.length) {
          await Promise.all(changedWatchedPeople.map(async (person) => {
            const title = t("presenceChangedTitle");
            const body = t("presenceChangedBody", {
              person: person.name,
              status: t("phoneFree")
            });
            await showPresenceNotification(title, body, person.id);
          }));
          const changedIds = new Set(changedWatchedPeople.map((person) => String(person.id)));
          setPresenceNotifications((current) =>
            current.filter((id) => !changedIds.has(String(id)))
          );
        }
      }
      previousPeopleRef.current = nextPeople;
      setPeople(nextPeople);
      setLastUpdated(data.generatedAt || new Date().toISOString());
      setError(null);
    } catch (e) {
      console.error(e);
      setError("apiLoadError");
    } finally {
      loadInProgressRef.current = false;
      setLoading(false);
      setRefreshing(false);
    }
  }, [presenceNotifications, showPresenceNotification, t, teamsContext.host]);

  useEffect(() => {
    load();
  }, [load]);
  useEffect(() => {
    const id = window.setInterval(() => load(), 15000);
    return () => window.clearInterval(id);
  }, [load]);

  const decorated = useMemo(
    () => people.map((p) => ({ ...p, favorite: favorites.includes(p.id) })),
    [people, favorites]
  );
  const scoped = useMemo(
    () =>
      view === "favorites"
        ? decorated.filter((p) => p.favorite)
        : view === "all"
          ? decorated
          : decorated.filter((p) => p.company === view),
    [decorated, view]
  );
  const departments = useMemo(
    () =>
      [...new Set(scoped.filter((p) => p.company !== MEETING_ROOM_COMPANY).map((p) => p.department))].sort((a, b) =>
        a.localeCompare(b, language)
      ),
    [scoped, language]
  );
  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return scoped.filter((p) => {
      const absent = Boolean(p.sageAbsence || p.oofAbsence);
      const departmentLabel =
        p.department === MISSING_DEPARTMENT ? t("departmentMissing") : p.department;
      const roleLabel = p.role === MISSING_ROLE ? t("roleMissing") : p.role;
      return (
        (department === "all" || p.department === department) &&
        (presence === "all" ||
          (presence === "absence" ? absent : !absent && p.presence === presence)) &&
        (phone === "all" || p.phone === phone) &&
        (location === "all" || p.location === location) &&
        (!q ||
          [p.name, roleLabel, departmentLabel, p.email, p.companyName]
            .filter(Boolean)
            .join(" ")
            .toLowerCase()
            .includes(q))
      );
    });
  }, [scoped, query, department, presence, phone, location, t]);
  const grouped = useMemo(
    () =>
      companies
        .map((company) => {
          const peopleInCompany = filtered.filter((person) => person.company === company.id);
          return {
            ...company,
            people: peopleInCompany,
            departments: company.id === MEETING_ROOM_COMPANY
              ? []
              : [...new Set(peopleInCompany.map((person) => person.department))]
                  .sort((left, right) => left.localeCompare(right, language))
                  .map((name) => ({
                    name,
                    people: peopleInCompany.filter((person) => person.department === name)
                  }))
          };
        })
        .filter((company) => company.id === MEETING_ROOM_COMPANY
          ? company.people.length > 0
          : company.departments.length > 0),
    [filtered, language]
  );
  const displayedGroups = useMemo(
    () => view === "all"
      ? [...grouped].sort((left, right) =>
          Number(right.id === MEETING_ROOM_COMPANY) - Number(left.id === MEETING_ROOM_COMPANY)
        )
      : grouped,
    [grouped, view]
  );
  const favoriteGroupSections = useMemo(() => {
    const assignments = favoriteGroupsData.assignments;
    const assignedGroupIds = new Set(favoriteGroupsData.groups.map((group) => group.id));
    const groups = [...favoriteGroupsData.groups].sort((left, right) =>
      left.name.localeCompare(right.name, language)
    );
    return [
      {
        id: "",
        name: t("ungroupedFavorites"),
        people: filtered.filter((person) => !assignedGroupIds.has(assignments[person.id]))
      },
      ...groups.map((group) => ({
        ...group,
        people: filtered.filter((person) => assignments[person.id] === group.id)
      }))
    ];
  }, [filtered, favoriteGroupsData, language, t]);
  const navCount = (id) =>
    id === "all"
      ? decorated.length
      : id === "favorites"
        ? decorated.filter((p) => p.favorite).length
        : decorated.filter((p) => p.company === id).length;
  const collapseAllSections = () => {
    if (view === "favorites") {
      setCollapsedFavoriteGroups((current) => ({
        ...current,
        ...Object.fromEntries(favoriteGroupSections.map((group) => [group.id, true]))
      }));
      return;
    }
    setCollapsedCompanies((current) => ({
      ...current,
      ...Object.fromEntries(displayedGroups.map((company) => [company.id, true]))
    }));
    setCollapsedDepartments((current) => ({
      ...current,
      ...Object.fromEntries(displayedGroups.flatMap((company) =>
        company.departments.map((group) => [`${company.id}:${group.name}`, true])
      ))
    }));
  };
  const expandAllSections = () => {
    if (view === "favorites") {
      setCollapsedFavoriteGroups((current) => ({
        ...current,
        ...Object.fromEntries(favoriteGroupSections.map((group) => [group.id, false]))
      }));
      return;
    }
    setCollapsedCompanies({});
    setCollapsedDepartments({});
  };
  const toggleFavorite = (id) => {
    const removingFavorite = favorites.includes(id);
    setFavorites((current) =>
      removingFavorite
        ? current.filter((favoriteId) => favoriteId !== id)
        : [...current, id]
    );
    if (removingFavorite) {
      setFavoriteGroupsData((groupData) => {
        const assignments = { ...groupData.assignments };
        delete assignments[id];
        return { ...groupData, assignments };
      });
    }
  };
  const togglePresenceNotification = (person) => {
    const personId = String(person.id);
    if (presenceNotifications.includes(personId)) {
      setPresenceNotifications((current) => current.filter((id) => String(id) !== personId));
      return;
    }
    setPresenceNotifications((current) =>
      current.includes(personId) ? current : [...current, personId]
    );
  };
  const createFavoriteGroup = (event) => {
    event.preventDefault();
    const name = newFavoriteGroupName.trim();
    if (!name) return;
    if (
      favoriteGroupsData.groups.some(
        (group) => group.name.toLocaleLowerCase(language) === name.toLocaleLowerCase(language)
      )
    ) {
      return;
    }
    const id = globalThis.crypto?.randomUUID?.() || `favorite-group-${Date.now()}`;
    setFavoriteGroupsData((current) => ({
      ...current,
      groups: [...current.groups, { id, name, color: "#64748B", icon: "folder" }]
    }));
    setNewFavoriteGroupName("");
  };
  const startEditingFavoriteGroup = (group) => {
    setCollapsedFavoriteGroups((current) => ({ ...current, [group.id]: false }));
    setEditingFavoriteGroupId(group.id);
    setFavoriteGroupDraft({
      name: group.name,
      color: group.color || "#64748B",
      icon: favoriteGroupIcons[group.icon] ? group.icon : "folder"
    });
  };
  const saveFavoriteGroup = (event, groupId) => {
    event.preventDefault();
    const name = favoriteGroupDraft?.name.trim();
    if (!name) return;
    setFavoriteGroupsData((current) => ({
      ...current,
      groups: current.groups.map((group) =>
        group.id === groupId ? { ...group, ...favoriteGroupDraft, name } : group
      )
    }));
    setEditingFavoriteGroupId(null);
    setFavoriteGroupDraft(null);
  };
  const deleteFavoriteGroup = (groupId) => {
    setFavoriteGroupsData((current) => {
      const assignments = Object.fromEntries(
        Object.entries(current.assignments).filter(([, assignedGroupId]) => assignedGroupId !== groupId)
      );
      return {
        groups: current.groups.filter((group) => group.id !== groupId),
        assignments
      };
    });
    if (editingFavoriteGroupId === groupId) {
      setEditingFavoriteGroupId(null);
      setFavoriteGroupDraft(null);
    }
  };
  const assignFavoriteGroup = (personId, groupId) => {
    setFavoriteGroupsData((current) => {
      const assignments = { ...current.assignments };
      if (groupId) assignments[personId] = groupId;
      else delete assignments[personId];
      return { ...current, assignments };
    });
  };
  const addFavoriteToGroup = (event, groupId) => {
    event.preventDefault();
    if (!favoriteToAdd) return;
    assignFavoriteGroup(favoriteToAdd, groupId);
    setFavoriteToAdd("");
    setAddingToFavoriteGroup(null);
  };
  const chooseView = (id) => {
    setView(id);
    setDepartment("all");
  };
  const chooseDefault = (id) => {
    setDefaultView(id);
    setView(id);
    setDepartment("all");
  };
  const reset = () => {
    setQuery("");
    setDepartment("all");
    setPresence("all");
    setPhone("all");
    setLocation("all");
  };

  if (loading && people.length === 0) {
    return (
      <div className={`app-shell ${theme}`}>
        <div className="loading-screen">
          <Logo />
          <strong>Crew Presence Board</strong>
          <span>{t("loading")}</span>
        </div>
      </div>
    );
  }

  return (
    <div className={`app-shell ${theme}`}>
      <header className="topbar">
        <Logo />
        <div className="brand-text">
          <strong>Crew Presence Board</strong>
          <span>{t("appSubtitle")}</span>
        </div>
        <div className="search-box">
          <Search size={16} />
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder={t("searchPlaceholder")}
          />
          {query && (
            <button onClick={() => setQuery("")} title={t("clearSearch")}>
              <X size={14} />
            </button>
          )}
        </div>
        <div className="header-actions">
          <button
            className="header-icon-button"
            onClick={() => setTheme((v) => (v === "dark" ? "light" : "dark"))}
            title={t("themeToggle")}
          >
            {theme === "dark" ? <Sun size={18} /> : <Moon size={18} />}
          </button>
          <button
            className="header-icon-button language-button"
            onClick={() => setLanguage((value) => (value === "de" ? "en" : "de"))}
            title={`${t("languageToggle")}: ${t(language === "de" ? "languageEnglish" : "languageGerman")}`}
            aria-label={`${t("languageToggle")}: ${language.toUpperCase()}`}
          >
            <Languages size={17} />
            <span>{language.toUpperCase()}</span>
          </button>
          <button
            className="header-icon-button"
            onClick={() => setInfoOpen(true)}
            title={t("information")}
          >
            <Info size={18} />
          </button>
          <button
            className="outline-button"
            onClick={() => load(true)}
            disabled={refreshing}
          >
            <RefreshCw size={15} className={refreshing ? "spin" : ""} />
            {t("refresh")}
            {lastUpdated && (
              <span className="refresh-time">
                {new Intl.DateTimeFormat(language === "de" ? "de-DE" : "en-GB", {
                  hour: "2-digit",
                  minute: "2-digit",
                  second: "2-digit"
                }).format(new Date(lastUpdated))}
              </span>
            )}
          </button>
        </div>
      </header>
      <main className="main-content">
        {error && (
          <div className="error-banner">
            <span>{t(error)}</span>
            <button onClick={() => load(true)}>{t("refreshAgain")}</button>
          </div>
        )}
        <div className="control-stack">
        <div className="toolbar">
          <nav className="company-nav">
            {[
              { id: "all", label: t("all"), icon: Building2 },
              { id: "favorites", label: t("favorites"), icon: Star },
              ...companies
                .filter((company) => company.id !== MISSING_COMPANY || navCount(MISSING_COMPANY) > 0)
                .map((c) => ({
                  id: c.id,
                  label: c.id === MISSING_COMPANY
                    ? t("companyMissing")
                    : c.id === MEETING_ROOM_COMPANY
                      ? t("meetingRoomCompany")
                      : c.id
                }))
            ].map((item) => {
              const Icon = item.icon;
              return (
                <button
                  key={item.id}
                  className={view === item.id ? "active" : ""}
                  onClick={() => chooseView(item.id)}
                >
                  {Icon && <Icon size={14} />} {item.label}
                  <small>{navCount(item.id)}</small>
                </button>
              );
            })}
          </nav>
          <button
            className="outline-button filter-button"
            onClick={() => setFiltersOpen((v) => !v)}
          >
            <SlidersHorizontal size={14} />
            {t("filter")}
            <ChevronDown size={13} className={filtersOpen ? "rotate" : ""} />
          </button>
          <select
            className="sort-select"
            aria-label={t("sortBy")}
            value={sortOrder}
            onChange={(event) => setSortOrder(event.target.value)}
          >
            <option value="name-asc">{t("sortNameAsc")}</option>
            <option value="name-desc">{t("sortNameDesc")}</option>
            <option value="availability">{t("sortAvailability")}</option>
          </select>
          <div className="accordion-actions">
            <button
              className="icon-button accordion-action-button"
              type="button"
              title={t("collapseAll")}
              aria-label={t("collapseAll")}
              onClick={collapseAllSections}
            >
              <ChevronsUp size={15} />
            </button>
            <button
              className="icon-button accordion-action-button"
              type="button"
              title={t("expandAll")}
              aria-label={t("expandAll")}
              onClick={expandAllSections}
            >
              <ChevronsDown size={15} />
            </button>
          </div>
        </div>
        {filtersOpen && (
          <div className="filter-panel">
            <select value={department} onChange={(e) => setDepartment(e.target.value)}>
              <option value="all">{t("allDepartments")}</option>
              {departments.map((v) => (
                <option key={v} value={v}>
                  {v === MISSING_DEPARTMENT ? t("departmentMissing") : v}
                </option>
              ))}
            </select>
            <select value={presence} onChange={(e) => setPresence(e.target.value)}>
              <option value="all">{t("allPresenceStatuses")}</option>
              <option value="absence">{t("absenceFilter")}</option>
              {Object.entries(presenceMap).map(([k, v]) => (
                <option key={k} value={k}>
                  {t(v.labelKey)}
                </option>
              ))}
            </select>
            <select value={phone} onChange={(e) => setPhone(e.target.value)}>
              <option value="all">{t("allPhoneStatuses")}</option>
              <option value="free">{t("phoneFreeFilter")}</option>
              <option value="call">{t("phoneCallFilter")}</option>
            </select>
            <select value={location} onChange={(e) => setLocation(e.target.value)}>
              <option value="all">{t("locationsBoth")}</option>
              <option value="office">{t("office")}</option>
              <option value="home">{t("homeOffice")}</option>
            </select>
            <button className="reset-button" onClick={reset}>
              <X size={14} />{t("reset")}
            </button>
          </div>
        )}
        </div>
        {view === "favorites" && (
          <section className="favorite-group-manager" aria-label={t("favoriteGroups")}>
            <form className="favorite-group-form" onSubmit={createFavoriteGroup}>
              <input
                value={newFavoriteGroupName}
                onChange={(event) => setNewFavoriteGroupName(event.target.value)}
                placeholder={t("favoriteGroupPlaceholder")}
                aria-label={t("favoriteGroupPlaceholder")}
                maxLength={40}
              />
              <button
                className="outline-button"
                type="submit"
                disabled={!newFavoriteGroupName.trim()}
              >
                <Plus size={14} />
                {t("createFavoriteGroup")}
              </button>
            </form>
          </section>
        )}
        <div className="result-count">
          <span>
            <strong>{filtered.length}</strong> {t("employeeLabel")}
          </span>
        </div>
        {view === "favorites"
          ? favoriteGroupSections.map((group) => (
              <section className="company-section" key={`favorites-${group.id}`}>
                <header>
                  <i style={{ backgroundColor: group.color || "#64748B" }} />
                  <button
                    className="favorite-group-collapse-toggle"
                    type="button"
                    aria-expanded={!collapsedFavoriteGroups[group.id]}
                    aria-controls={`favorite-group-content-${group.id || "ungrouped"}`}
                    title={t(
                      collapsedFavoriteGroups[group.id]
                        ? "expandFavoriteGroup"
                        : "collapseFavoriteGroup"
                    )}
                    onClick={() =>
                      setCollapsedFavoriteGroups((current) => ({
                        ...current,
                        [group.id]: !current[group.id]
                      }))
                    }
                  >
                    <ChevronDown
                      size={14}
                      className={collapsedFavoriteGroups[group.id] ? "rotate" : ""}
                    />
                    <FavoriteGroupIcon name={group.icon} color={group.color} />
                    <h2>{group.name}</h2>
                    <span>({group.people.length})</span>
                  </button>
                  <button
                    className="favorite-group-add-button add"
                    type="button"
                    title={`${t("addFavoriteToGroup")}: ${group.name}`}
                    aria-label={`${t("addFavoriteToGroup")}: ${group.name}`}
                    onClick={() => {
                      setAddingToFavoriteGroup((current) =>
                        current === group.id ? null : group.id
                      );
                      setFavoriteToAdd("");
                    }}
                  >
                    <Plus size={14} />
                  </button>
                  {group.id && (
                    <>
                      <button
                        className="favorite-group-add-button edit"
                        type="button"
                        title={`${t("editFavoriteGroup")}: ${group.name}`}
                        aria-label={`${t("editFavoriteGroup")}: ${group.name}`}
                        onClick={() => {
                          if (editingFavoriteGroupId === group.id) {
                            setEditingFavoriteGroupId(null);
                            setFavoriteGroupDraft(null);
                          } else {
                            startEditingFavoriteGroup(group);
                          }
                        }}
                      >
                        <Pencil size={13} />
                      </button>
                      <button
                        className="favorite-group-add-button remove"
                        type="button"
                        title={`${t("deleteFavoriteGroup")}: ${group.name}`}
                        aria-label={`${t("deleteFavoriteGroup")}: ${group.name}`}
                        onClick={() => deleteFavoriteGroup(group.id)}
                      >
                        <Trash2 size={13} />
                      </button>
                    </>
                  )}
                </header>
                <div
                  id={`favorite-group-content-${group.id || "ungrouped"}`}
                  hidden={collapsedFavoriteGroups[group.id]}
                >
                    {editingFavoriteGroupId === group.id && favoriteGroupDraft && (
                      <form
                        className="favorite-group-settings"
                        onSubmit={(event) => saveFavoriteGroup(event, group.id)}
                      >
                        <label>
                          <span>{t("favoriteGroupName")}</span>
                          <input
                            value={favoriteGroupDraft.name}
                            onChange={(event) =>
                              setFavoriteGroupDraft((current) => ({
                                ...current,
                                name: event.target.value
                              }))
                            }
                            maxLength={40}
                          />
                        </label>
                        <label>
                          <span>{t("favoriteGroupColor")}</span>
                          <input
                            type="color"
                            value={favoriteGroupDraft.color}
                            onChange={(event) =>
                              setFavoriteGroupDraft((current) => ({
                                ...current,
                                color: event.target.value
                              }))
                            }
                          />
                        </label>
                        <label>
                          <span>{t("favoriteGroupIcon")}</span>
                          <select
                            value={favoriteGroupDraft.icon}
                            onChange={(event) =>
                              setFavoriteGroupDraft((current) => ({
                                ...current,
                                icon: event.target.value
                              }))
                            }
                          >
                            {favoriteGroupIconOptions.map((option) => (
                              <option key={option.id} value={option.id}>
                                {t(option.labelKey)}
                              </option>
                            ))}
                          </select>
                        </label>
                        <button
                          className="outline-button"
                          type="submit"
                          disabled={!favoriteGroupDraft.name.trim()}
                        >
                          {t("save")}
                        </button>
                        <button
                          className="outline-button"
                          type="button"
                          onClick={() => {
                            setEditingFavoriteGroupId(null);
                            setFavoriteGroupDraft(null);
                          }}
                        >
                          {t("cancel")}
                        </button>
                      </form>
                    )}
                    {addingToFavoriteGroup === group.id && (
                      <form
                        className="favorite-add-form"
                        onSubmit={(event) => addFavoriteToGroup(event, group.id)}
                      >
                        <select
                          value={favoriteToAdd}
                          onChange={(event) => setFavoriteToAdd(event.target.value)}
                          aria-label={t("chooseFavoriteToAdd")}
                        >
                          <option value="">{t("chooseFavoriteToAdd")}</option>
                          {sortPeople(
                            decorated.filter(
                              (person) =>
                                person.favorite &&
                                (favoriteGroupsData.assignments[person.id] || "") !== group.id
                            ),
                            "name-asc",
                            language
                          ).map((person) => (
                            <option key={person.id} value={person.id}>
                              {person.name}
                            </option>
                          ))}
                        </select>
                        <button
                          className="outline-button"
                          type="submit"
                          disabled={!favoriteToAdd}
                        >
                          {t("addFavorite")}
                        </button>
                        <button
                          className="icon-button"
                          type="button"
                          title={t("cancel")}
                          aria-label={t("cancel")}
                          onClick={() => {
                            setAddingToFavoriteGroup(null);
                            setFavoriteToAdd("");
                          }}
                        >
                          <X size={13} />
                        </button>
                      </form>
                    )}
                    <div className="people-grid">
                      {sortPeople(group.people, sortOrder, language).map((person) =>
                        person.company === MEETING_ROOM_COMPANY ? (
                          <ResourceCard
                            key={person.id}
                            person={person}
                            onOpenCalendar={setResourceCalendarPerson}
                            onOpenInfo={setResourceInfoPerson}
                            t={t}
                            language={language}
                          />
                        ) : (
                          <PersonCard
                            key={person.id}
                            person={person}
                            company={companies.find((item) => item.id === person.company) || companies[0]}
                            favorite={person.favorite}
                            onFavorite={toggleFavorite}
                            watchingPresence={presenceNotifications.includes(String(person.id))}
                            onWatchPresence={togglePresenceNotification}
                            t={t}
                          />
                        )
                      )}
                    </div>
                </div>
              </section>
            ))
          : displayedGroups.map((company) => {
              const companyCollapsed = Boolean(collapsedCompanies[company.id]);
              const companyContentId = `company-content-${company.id}`;
              const companyName = company.id === MISSING_COMPANY
                ? t("companyMissing")
                : company.id === MEETING_ROOM_COMPANY
                  ? t("meetingRoomCompany")
                  : company.name;

              return (
                <section className="company-section" key={company.id}>
                  <header>
                    <i style={{ backgroundColor: company.color }} />
                    <button
                      className="company-collapse-toggle"
                      type="button"
                      aria-expanded={!companyCollapsed}
                      aria-controls={companyContentId}
                      title={t(companyCollapsed ? "expandCompany" : "collapseCompany")}
                      onClick={() => setCollapsedCompanies((current) => ({
                        ...current,
                        [company.id]: !current[company.id]
                      }))}
                    >
                      <ChevronDown className={`collapse-chevron${companyCollapsed ? " is-collapsed" : ""}`} size={15} />
                      <h2>{companyName}</h2>
                    </button>
                  </header>
                  <div id={companyContentId} hidden={companyCollapsed}>
                    {company.id === MEETING_ROOM_COMPANY ? (
                      <div className="people-grid">
                        {sortPeople(company.people, sortOrder, language).map((person) => (
                          <ResourceCard
                            key={person.id}
                            person={person}
                            onOpenCalendar={setResourceCalendarPerson}
                            onOpenInfo={setResourceInfoPerson}
                            t={t}
                            language={language}
                          />
                        ))}
                      </div>
                    ) : company.departments.map((group) => {
                      const departmentKey = `${company.id}:${group.name}`;
                      const departmentCollapsed = Boolean(collapsedDepartments[departmentKey]);
                      const departmentContentId = `department-content-${company.id}-${encodeURIComponent(group.name)}`;
                      const departmentName = group.name === MISSING_DEPARTMENT
                        ? t("departmentMissing")
                        : group.name;
                      const departmentHeading = (
                        <>
                          {departmentName} <span>({group.people.length})</span>
                        </>
                      );

                      return (
                        <div className="department" key={departmentKey}>
                          <button
                            className="department-collapse-toggle"
                            type="button"
                            aria-expanded={!departmentCollapsed}
                            aria-controls={departmentContentId}
                            title={t(departmentCollapsed ? "expandDepartment" : "collapseDepartment")}
                            onClick={() => setCollapsedDepartments((current) => ({
                              ...current,
                              [departmentKey]: !current[departmentKey]
                            }))}
                          >
                            <ChevronDown className={`collapse-chevron${departmentCollapsed ? " is-collapsed" : ""}`} size={14} />
                            <h3>{departmentHeading}</h3>
                          </button>
                          <div id={departmentContentId} hidden={departmentCollapsed}>
                            <div className="people-grid">
                              {sortPeople(group.people, sortOrder, language).map((person) => (
                                <PersonCard
                                  key={person.id}
                                  person={person}
                                  company={company}
                                  favorite={person.favorite}
                                  onFavorite={toggleFavorite}
                                  watchingPresence={presenceNotifications.includes(String(person.id))}
                                  onWatchPresence={togglePresenceNotification}
                                  t={t}
                                />
                              ))}
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </section>
              );
            })}
        {(view === "favorites"
          ? filtered.length === 0
          : grouped.length === 0) && (
          <div className="empty-state">
            <Users size={36} />
            <strong>{t("noPeople")}</strong>
            <button className="outline-button" onClick={reset}>
              {t("resetFilters")}
            </button>
          </div>
        )}
      </main>
      {notificationToast && (
        <div className="notification-toast" role="status" aria-live="polite">
          <div>
            <strong>{notificationToast.title}</strong>
            <p>{notificationToast.body}</p>
          </div>
          <button
            className="icon-button"
            type="button"
            title={t("close")}
            aria-label={t("close")}
            onClick={() => {
              window.clearTimeout(notificationToastTimeoutRef.current);
              setNotificationToast(null);
            }}
          >
            <X size={15} />
          </button>
        </div>
      )}
      <InfoModal
        open={infoOpen}
        onClose={() => setInfoOpen(false)}
        defaultView={defaultView}
        onDefaultView={chooseDefault}
        t={t}
      />
      <ResourceCalendarModal
        person={resourceCalendarPerson}
        onClose={() => setResourceCalendarPerson(null)}
        t={t}
        language={language}
      />
      <ResourceInfoModal
        key={resourceInfoPerson?.id || "closed"}
        person={resourceInfoPerson}
        onClose={() => setResourceInfoPerson(null)}
        t={t}
      />
    </div>
  );
}
