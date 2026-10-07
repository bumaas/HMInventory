# HMInventory — Projektwissen

Symcon-Modulbibliothek (`bumaas`) mit genau einem Modul: **HM Inventory Report Creator**
(`HM_Inventory/`, Typ Device, Prefix `HMI`). Erstellt per XML-RPC einen HTML-Report aller
HomeMatic-Geräte einer CCU (BidCos-RF, HmIP, BidCos-Wired) inkl. RSSI-Werten.

## Struktur

- `HM_Inventory/module.php` — Klasse `HMInventoryReportCreator extends IPSModuleStrict` (strict_types).
  Ablauf `CreateReport()`: `collectDeviceData()` (XML-RPC: `listDevices` je Service RF/IP/WR,
  `listBidcosInterfaces`, `rssiInfo`) → Render-Methoden (`renderHeaderSection` …) → HTML-Datei.
- `libs/phpxmlrpc/` — vendorte [phpxmlrpc](https://github.com/gggeek/phpxmlrpc) 4.11.5,
  **nur `src/`** + Lizenz/README (keine Tests/Demos/Debugger einchecken — Symcon scannt den
  Modulbaum). Einbindung über `src/Autoloader.php`, genutzt werden `Client`, `Request`,
  `Response`, `Encoder`. `Client::setSSLVerify*` ist deprecated → `setOption(Client::OPT_VERIFY_*)`.
- `tests/check_locale.php` — Übersetzungs-Vollständigkeitscheck (Muster aus BlindControl),
  läuft in der CI (`.github/workflows/check.yml`, PHP 8.4: Code-Stil mit php-cs-fixer gegen
  das Regelwerk im Submodul `.style` — `libs/phpxmlrpc` per `.style-exclude` ausgenommen —,
  php -l über `HM_Inventory/`, `tests/` und `libs/`, JSON-Validität, Locale-Check, danach
  per Glob jede `tests/check-*.php`).
- `tests/check-readme.php` — Doku-Sperrklinke: Module, Formularfelder (Name/Beschriftung als
  Stichwort, also erste Tabellenspalte, Überschrift, Listenpunkt oder Fettdruck) und jede
  öffentliche `HMI_*`-Funktion müssen im README stehen; kein „IP-Symcon“ in Doku, `form.json`,
  `locale.json`, `module.php` (URLs ausgenommen). `tests/readme-bekannt.json` ist leer und
  **bleibt leer**: Eine neue Lücke wird dokumentiert, nicht eingetragen (`--bekannt-schreiben`
  streicht nur und bricht bei neuen Lücken mit Exit 1 ab). Die Prüfung selbst ist durch
  `tests/check-readme-selbsttest.php` abgesichert (echtes README, gezielt verändert).
- `tests/check-report-logik.php` — Report-Logik gegen den Kernel-Stub (`tests/harness.php`,
  Submodul `tests/stubs` = symcon/SymconStubs, gepinnt; in CI aus `php -l`, JSON-Prüfung und
  php-cs-fixer ausgenommen): Sortierung, gelbe Markierung (`markBestInterface`), Beschriftung.
  Fixture `tests/fixtures/devicelist_nuc_2026-09-30.json` ist die echte `DeviceList` der
  nuc-Instanz #10064 (362 Einträge, Namen geleert). Private Methoden ruft der Harness per
  Reflection (`rufe`, `rufeMitReferenz`).

## Besonderheiten / Stolpersteine

- Der HM-Script-Zugriff (`getHMChannelName` → `SendScript` → cURL auf `Script.exe` der CCU)
  ist ein separater Transportweg neben XML-RPC; Fehler dort werden geschluckt (`false` → Name '').
- Die Variable `DeviceList` (Ident) enthält die Geräteliste **JSON-kodiert** (kein HTML);
  registriert mit `VARIABLE_PRESENTATION_VALUE_PRESENTATION`. Schlüsselnamen der JSON-Einträge
  (`IPS_occ`, `HM_address`, …) sind Datenkontrakt — nicht umbenennen.
- Der Report-Link (`buildUserReportLink`) funktioniert nur für Ausgabedateien unterhalb
  `<kernel>/user/`; Schema/Port kommen aus der ersten aktiven WebServer-Instanz
  (GUID `{D83E9CCF-9869-420F-8306-2B043E9BA180}`), Fallback ist die fest eingebaute
  Weboberfläche auf Port `3777`.
- Der Report ist über `Translate()` beschriftet; die `locale.json` trägt zwei Schlüssel, die
  keine Texte sind: `"en": "de"` liefert das `lang`-Attribut des Reports (Sprachcode aus der
  Übersetzung, nicht aus `IPS_GetSystemLanguage()`), `"Y-m-d H:i:s": "d.m.Y H:i:s"` das
  Datumsformat der Kopfzeile. Eine neue Sprache braucht beide. Die Roaming-Kopfzelle heißt
  `Roa&shy;ming` (weicher Trennstrich, 2%-Spalte), kein `overflow-wrap`.
  `tests/check-report-uebersetzung.php` prüft das mit umschaltbarer Harness-Sprache
  (`HMInventoryHarness::$sprache`); Interfaces aus der Fixture liefert `fixtureInterfaces()`
  (filtert die leeren `HM_Interface`-Werte der HmIP-Einträge).
- `compareByAddress` gibt bei gleichem Geräteteil `0/1` (nie `-1`) zurück — historisches
  Verhalten, bei Refactorings beibehalten (Sortierstabilität des Reports).
- Timer `Update` ruft `IPS_RequestAction(…, 'CreateReport', true)`; die öffentliche API
  (`HMI_CreateReport`, `HMI_GetReportUrl`, `HMI_GetOutputFileAbsolutePath`) ist im README dokumentiert.

## Versionspflege

Siehe globale `CLAUDE.md`, Abschnitt „Symcon: Build-/Versionspflege in Modul-Repos".
