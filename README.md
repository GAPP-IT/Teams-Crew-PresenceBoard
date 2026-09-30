# Crew Presence Board

Vollständiges React-Frontend auf Basis von Vite. Enthalten sind Firmen-Navigation, Favoriten, Suche, Filter, kompakte Personenkarten, Teams-Präsenz, separater Telefoniestatus, Standort, Termine und priorisierte Abwesenheit.

## Voraussetzungen

- Node.js 20 oder neuer empfohlen
- npm

## Installation unter Windows

```powershell
cd C:\NODEJS
mkdir crew-presence-board
cd crew-presence-board
```

ZIP-Inhalt in diesen Ordner entpacken. Danach:

```powershell
npm install
npm run dev
```

Im Browser öffnen:

```text
http://localhost:5173
```

## Produktions-Build

```powershell
npm run build
```

Die fertigen Dateien liegen danach in `dist`.

## Schnell austauschbare Dateien

- `src/App.jsx`: Oberfläche und Anwendungslogik
- `src/styles.css`: komplettes Design
- `src/data/mockPeople.js`: Firmen und Testpersonen
- `.env`: spätere URL des Node-/Graph-Backends

## Spätere API-Anbindung

In `App.jsx` ist im `refresh()`-Handler bereits markiert, wo der Fetch-Aufruf ergänzt wird. Erwartetes Personenmodell:

```json
{
  "id": "entra-object-id",
  "name": "Max Mustermann",
  "initials": "MM",
  "company": "PAG",
  "department": "IT & Infrastructure",
  "role": "Systemadministrator",
  "email": "max@example.com",
  "presence": "available",
  "phone": "free",
  "location": "office",
  "currentMeetingEnd": null,
  "nextMeeting": "13:30",
  "sageAbsence": null,
  "oofAbsence": null,
  "favorite": false
}
```

Zulässige Werte:

- `presence`: `available`, `busy`, `meeting`, `dnd`, `away`, `offline`
- `phone`: `free`, `call`, `ringing`, `unavailable`
- `location`: `office`, `home`

Abwesenheits-Priorität im Frontend:

1. `sageAbsence`
2. `oofAbsence`
3. normaler Teams-Status

Der Abwesenheitsgrund wird nicht angezeigt. Bei „Außer Haus“ hat die Abwesenheitsanzeige Vorrang vor dem Teams-Status und Terminen; angezeigt wird „Zurück am DD.MM.YYYY“.

Für die Rückkehrzeit aus Outlook-Antwortregeln benötigt die Entra-App die Microsoft-Graph-Anwendungsberechtigung `MailboxSettings.Read` mit Admin-Zustimmung. OOF-Kalendereinträge werden weiterhin als Fallback ausgewertet.
