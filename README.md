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

## Benachrichtigungen

Die Glocke speichert Beobachtungen lokal im Browser. Bei erlaubten Browser-Systemmeldungen werden Statuswechsel auch außerhalb des Tabs angezeigt; andernfalls erscheint ein Hinweis im geöffneten Board. Für Meldungen bei geschlossenem Tab wäre eine separate serverseitige Teams-Integration nötig.

Um die frühere Anmeldung auch aus einer bereits installierten Teams-App zu entfernen, `Teams App/Teams-App-1.0.5.zip` in Teams als App-Update hochladen und installieren.

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

## Teams-Hintergrundbenachrichtigungen

Wird die Glocke in der installierten Teams-App aktiviert, speichert die API das Abo pro Teams-Benutzer. Der dauerhaft laufende API-Prozess prüft die Präsenz standardmäßig alle 30 Sekunden. Beim Wechsel von „Telefoniert“ oder „Abwesend“ zu „Verfügbar“ sendet er eine Teams-Aktivitätsmeldung und entfernt das Abo als Einmalmeldung. Teams zeigt den Eintrag im Aktivitätsfeed; ob zusätzlich ein Banner/Toast erscheint, hängt von den Teams-Benachrichtigungseinstellungen des Empfängers ab.

Die Funktion benötigt einmalige Entra-/Teams-Admin-Konfiguration:

1. In der Entra-App unter **Eine API verfügbar machen** eine Application ID URI festlegen, einen delegierten Scope `access_as_user` anlegen und die Teams-Clientanwendungen als autorisierte Clients eintragen.
2. `webApplicationInfo.id` und `webApplicationInfo.resource` in `Teams App/manifest.json` müssen zu dieser Entra-App und Application ID URI passen. Teams SSO muss für die statische Registerkarte aktiviert sein.
3. Der Entra-App die Microsoft-Graph-Anwendungsberechtigung `TeamsActivity.Send` hinzufügen und Admin-Zustimmung erteilen. Die Teams-App muss bei jedem Empfänger installiert sein.
4. In `.env` die Werte passend zur öffentlichen HTTPS-URL setzen:

  ```env
  TEAMS_SSO_RESOURCE=api://<host>/<client-id>
  TEAMS_APP_URL=https://<host>
  TEAMS_NOTIFICATIONS_STORE_PATH=data/teams-notifications.json
  TEAMS_NOTIFICATION_POLL_INTERVAL_MS=30000
  ```

  `TEAMS_SSO_RESOURCE` muss exakt `webApplicationInfo.resource` entsprechen. `TENANT_ID`, `CLIENT_ID` und `CLIENT_SECRET` der Entra-App werden ebenfalls serverseitig benötigt.
5. Bei einer Domainänderung die URLs in Manifest, Entra-App und `.env` gemeinsam aktualisieren. Danach `Teams App/Teams-App-1.0.6.zip` als App-Update in Teams hochladen und installieren; eine reine Website-Aktualisierung übernimmt keine Manifeständerungen.

Der API-Prozess muss dauerhaft laufen. `TEAMS_NOTIFICATIONS_STORE_PATH` muss auf dauerhaft beschreibbaren Speicher zeigen. Bei mehreren API-Instanzen darf nur eine den Präsenz-Poller ausführen, damit keine doppelten Aktivitätsmeldungen entstehen.

Zur Diagnose zeigt `/api/health` unter `teamsNotifications` `configured`, `activeSubscriptions`, den letzten Prüfzeitpunkt und den letzten Graph-Sendeversuch. Nach Aktivieren einer Glocke muss `activeSubscriptions` steigen; nach einer zugestellten Einmalmeldung sinkt der Wert wieder. Graph-Ablehnungen erscheinen in `lastNotificationError`.
