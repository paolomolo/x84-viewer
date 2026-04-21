# x84-viewer

Lokaler Web-Viewer fuer GAEB DA XML (`.x84`) Dateien mit Tabellenansicht, XML-Baum und Excel-Export.

## Lokal starten

- Repository klonen
- `index.html` im Browser oeffnen

Die App ist komplett clientseitig und benoetigt keinen Server.

## Deployment im Web (GitHub Pages)

Dieses Repository enthaelt einen Workflow unter `.github/workflows/deploy-pages.yml`, der bei Push auf `main` automatisch auf GitHub Pages deployed.

Einmalig im GitHub-Repository einstellen:

1. **Settings -> Pages** oeffnen
2. Unter **Build and deployment** die Quelle auf **GitHub Actions** setzen
3. Aenderung speichern

Danach wird bei jedem Push auf `main` automatisch deployed.
