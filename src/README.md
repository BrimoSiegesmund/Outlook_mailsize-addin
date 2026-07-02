# Mailgröße Anzeige – Outlook Add-in

Zeigt beim Verfassen einer neuen E-Mail live die aktuelle Größe (Body + Anhänge)
an, inkl. einer Schätzung der tatsächlichen Sendegröße (Base64-Overhead) und
einer Warnung bei Überschreitung eines Schwellenwerts (Standard: 20 MB).

Funktioniert im **neuen Outlook**, in **Outlook on the web (OWA)** und im
**klassischen Outlook (Win32)** – jeweils als sogenanntes "Office Add-in"
(Office.js-basiert, kein VBA/COM-Add-in nötig).

## Wichtig: Hosting erforderlich

Outlook-Add-ins laden ihre Oberfläche (HTML/JS/CSS) von einer **HTTPS-URL**.
Es reicht nicht, die Dateien nur lokal zu haben – sie müssen irgendwo per
HTTPS erreichbar sein, damit Outlook sie laden kann. Optionen, von einfach
bis professionell:

1. **GitHub Pages** (kostenlos, am einfachsten):
   - Repo erstellen, den Ordner `src/` + `manifest.xml` hochladen
   - Pages in den Repo-Settings aktivieren
   - Alle `https://DEIN-HOSTING-DOMAIN` Platzhalter in `manifest.xml` durch
     deine GitHub-Pages-URL ersetzen, z. B.
     `https://timo-username.github.io/mailgroesse-addin`

2. **Azure Static Web Apps** (kostenlose Stufe verfügbar) – bietet mehr
   Kontrolle, falls ihr das intern bei Mels&Back sauber einbinden wollt.

3. **Firmeninterner Webserver** (z. B. IIS) mit gültigem SSL-Zertifikat,
   falls ihr es intern für alle 30 Mitarbeiter ausrollen wollt.

Danach in `manifest.xml` **alle** Vorkommen von `DEIN-HOSTING-DOMAIN` durch
die echte Domain ersetzen (auch das `AppDomains`-Element!).

## Installation zum Testen (Sideloading)

**Neues Outlook / Outlook on the web:**
1. Outlook öffnen → Einstellungen (Zahnrad) → *Add-Ins* → *Meine Add-Ins*
2. Unten "Benutzerdefiniertes Add-In hinzufügen" → "Aus Datei hinzufügen"
3. Die angepasste `manifest.xml` auswählen

**Klassisches Outlook (Win32):**
1. Datei → Add-Ins verwalten (öffnet OWA-Add-In-Verwaltung) → wie oben, oder
2. Über *Einfügen → Add-Ins abrufen → Eigene Add-Ins → Von Datei hinzufügen*

Nach der Installation erscheint beim Verfassen einer neuen Mail im Menüband
ein Button **"Größe anzeigen"**, der das Panel mit der Live-Anzeige öffnet.

## Anpassbare Werte (in `src/taskpane.js`)

```js
const WARN_THRESHOLD_MB = 20;   // Warnschwelle, an euer M365-Limit anpassen
const BASE64_OVERHEAD = 1.37;   // Kodierungs-Overhead für Anhänge
```

Das tatsächliche Anhang-Limit ist tenant-abhängig (M365-Standard: 20 MB,
von Admins bis 150 MB erhöhbar). Passt `WARN_THRESHOLD_MB` an euer
tatsächliches Limit an, falls euer Exchange-Admin es geändert hat.

## Bekannte Einschränkungen

- Die Body-Größe wird als UTF-8-Byte-Länge des HTML-Inhalts berechnet – das
  ist eine sehr gute Näherung, aber nicht exakt byteidentisch mit dem später
  tatsächlich gesendeten MIME-Teil (Quoted-Printable-Encoding kann nochmal
  leicht abweichen, üblicherweise < 5 %).
- Inline-Bilder im Body zählen als Anhänge (`isInline: true`) und werden in
  der Liste mit angezeigt – falls gewünscht, könnt ihr das im JS filtern.
- Für Signaturen, die automatisch von Outlook eingefügt werden, gilt: Diese
  sind im Body bereits enthalten und werden mitgezählt.
