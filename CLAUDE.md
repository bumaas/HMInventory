# HMInventory — Projektwissen

IP-Symcon-Modulbibliothek (`bumaas`) mit genau einem Modul: **HM Inventory Report Creator**
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
  läuft in der CI (`.github/workflows/check.yml`: php -l, JSON-Validität, Locale-Check).

## Besonderheiten / Stolpersteine

- Der HM-Script-Zugriff (`getHMChannelName` → `SendScript` → cURL auf `Script.exe` der CCU)
  ist ein separater Transportweg neben XML-RPC; Fehler dort werden geschluckt (`false` → Name '').
- Die Variable `DeviceList` (Ident) enthält die Geräteliste **JSON-kodiert** (kein HTML);
  registriert mit `VARIABLE_PRESENTATION_VALUE_PRESENTATION`. Schlüsselnamen der JSON-Einträge
  (`IPS_occ`, `HM_address`, …) sind Datenkontrakt — nicht umbenennen.
- Der Report-Link (`buildUserReportLink`) funktioniert nur für Ausgabedateien unterhalb
  `<kernel>/user/`; Webserver-Port ist fest `82`.
- `compareByAddress` gibt bei gleichem Geräteteil `0/1` (nie `-1`) zurück — historisches
  Verhalten, bei Refactorings beibehalten (Sortierstabilität des Reports).
- Timer `Update` ruft `IPS_RequestAction(…, 'CreateReport', true)`; die öffentliche API
  (`HMI_CreateReport`, `HMI_GetReportUrl`, `HMI_GetOutputFileAbsolutePath`) ist im README dokumentiert.

## Versionspflege

`library.json`: bei jeder inhaltlichen Änderung `build` +1 und `date` aktualisieren;
Commit-Subject: `<version> build <NN>: <Beschreibung>`.
