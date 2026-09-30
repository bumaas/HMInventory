# HM Inventory

[![Checks](https://github.com/bumaas/HMInventory/actions/workflows/check.yml/badge.svg)](https://github.com/bumaas/HMInventory/actions/workflows/check.yml)

Erstellt eine Übersicht aller Homematic-Geräte einer CCU als HTML-Seite: jeder Kanal mit Adresse, Gerätetyp, Firmware und der zugehörigen Symcon-Instanz, dazu die Funkpegel zwischen den Geräten und jedem Funk-Interface der CCU. Auf Wunsch landet dieselbe Liste als JSON in einer Variable.

[English version](README.en.md)

### Inhalt

1. [Wann brauche ich das?](#1-wann-brauche-ich-das)
2. [Installation und erster Lauf](#2-installation-und-erster-lauf)
3. [Den Report lesen](#3-den-report-lesen)
4. [Statusvariable](#4-statusvariable)
5. [Automatisieren](#5-automatisieren)
6. [Funktionen für Skripte](#6-funktionen-für-skripte)
7. [Grenzen](#7-grenzen)
8. [Begriffe](#8-begriffe)
9. [Anhang: Technik](#9-anhang-technik)

## 1. Wann brauche ich das?

- **Ein Gerät reagiert unzuverlässig.** Der Report zeigt für jedes Homematic-Funkgerät (BidCos) den Pegel zu jedem Interface der CCU und markiert das zugeordnete. Schwache Verbindungen fallen so auf einen Blick auf, ebenso ein Gerät, das an einem ungünstigen Interface hängt.
- **Was ist angelernt, aber nicht in Symcon angelegt?** Kanäle, die in der CCU existieren, aber keine Symcon-Instanz haben, erscheinen mit „-“ statt einer Objekt-ID.
- **Doppelt angelegte Instanzen finden.** Hängen zwei Symcon-Instanzen am selben Homematic-Kanal, stehen beide in roter Schrift.
- **Bestandsliste.** Welche Geräte mit welcher Firmware gibt es, und wie heißen sie in der CCU und in Symcon?

## 2. Installation und erster Lauf

**Voraussetzungen**

- Symcon ab Version 9.0
- Eine CCU (CCU2, CCU3, RaspberryMatic o. ä.), die in Symcon über einen **HomeMatic Socket** eingebunden ist. Das Modul nutzt dessen Adresse, Ports, SSL-Einstellung und Zugangsdaten und braucht keine eigenen.

**In Symcon**

1. Das Modul im Module Store unter *Homematic Inventory Module* installieren.
2. An beliebiger Stelle eine Instanz *HM Inventory Report Creator* anlegen. Sie hängt sich an den HomeMatic Socket. Bei mehreren CCUs je CCU eine Instanz anlegen und über *Gateway ändern* dem richtigen Socket zuordnen.
3. Einstellungen prüfen (Name im Formular in Klammern):

| Feld | Standard | Bedeutung |
| :--- | :------- | :-------- |
| `active` (aktiv) | an | Eine inaktive Instanz erstellt keinen Report. |
| `OutputFile` (Ausgabedatei) | `<Symcon-Verzeichnis>/user/HM_inventory.html` | Pfad und Name der HTML-Datei. Relative Angaben gelten ab dem Symcon-Verzeichnis. Nur unterhalb von `user/` ist der Report im Browser erreichbar. |
| `SortOrder` (Sortierung) | HM Adresse | Reihenfolge der Zeilen: HM Adresse, HM Gerätetyp, HM Kanaltyp, HM Gerätename oder IPS Gerätename |
| `ShowLongIPSDeviceNames` (Gebe lange IPS Gerätenamen aus) | aus | Zeigt den Namen der Symcon-Instanz mit vollem Pfad im Objektbaum. |
| `ShowHMConfiguratorDeviceNames` (Gebe HM Gerätenamen aus) | an | Zusätzliche Spalte mit dem Kanalnamen aus der CCU. Kostet Zeit, siehe [Grenzen](#7-grenzen). |
| `ShowNotUsedChannels` (Gebe ungenutzte Kanäle aus) | an | Listet auch die Kanäle, zu denen es keine Symcon-Instanz gibt. |
| `ShowMaintenanceEntries` (Gebe Maintenance Kanäle aus) | an | Nimmt dabei die Wartungskanäle (Kanal 0) mit. Wirkt nur, wenn ungenutzte Kanäle angezeigt werden. |
| `ShowVirtualKeyEntries` (Gebe Virtual Keys Einträge aus) | aus | Nimmt dabei die virtuellen Tasten der CCU mit. Wirkt nur, wenn ungenutzte Kanäle angezeigt werden. |
| `SaveDeviceListInVariable` (Sichere Geräteliste in einer Variablen) | aus | Legt die Variable *Geräteliste* an und schreibt die Liste bei jedem Report als JSON hinein, siehe [Statusvariable](#4-statusvariable). |
| `UpdateInterval` (Aktualisierungsintervall) | 0 | Erstellt den Report alle n Minuten neu. 0 = nur auf Knopfdruck oder per Skript. |

4. Übernehmen und **Erzeuge Bericht** drücken. Ein Fortschrittsbalken läuft durch, am Ende meldet das Formular „Report erstellt:“ mit dem Pfad der Datei. Liegt die Datei unterhalb von `user/`, erscheint der Knopf **Report im Browser öffnen**. Er zeigt die Adresse des Reports an.

**Meldungen der Instanz**

| Status | Bedeutung | Was tun? |
| :----- | :-------- | :------- |
| aktiv | bereit | – |
| inaktiv | Die Instanz ist abgeschaltet oder der HomeMatic Socket ist nicht aktiv. | `active` einschalten bzw. den Socket prüfen |
| Die angegebene Datei ist nicht erstellbar | Die Ausgabedatei ist leer, ihr Verzeichnis fehlt oder ist nicht beschreibbar. | Pfad korrigieren oder Verzeichnis anlegen |

**Meldungen beim Erstellen**

| Meldung | Bedeutung |
| :------ | :-------- |
| Es ist kein Gateway konfiguriert! | Die Instanz hängt an keinem HomeMatic Socket. |
| Die Instanz ist nicht aktiv! | wie Status *inaktiv* |
| Die Datei „…“ ist nicht schreibbar! | Die Datei ließ sich nicht schreiben, z. B. weil sie gerade gesperrt ist. |
| Fehler! | Der Report wurde nicht erstellt. Im Symcon-Log steht der Grund, etwa `Can't get any device information from the BidCos-Services`: Die CCU hat auf keinem Dienst geantwortet. |

## 3. Den Report lesen

Der Report ist englisch beschriftet. Er beginnt mit einer **Zusammenfassung**: die Zahl der Funk-Interfaces (und wie viele davon verbunden sind), die Zahl der Geräte je Dienst (HM-RF, HM-Wired, HmIP) und die Zahl der Symcon-Instanzen, die an diesem HomeMatic Socket hängen. Rechts daneben stehen die **Funk-Interfaces** der CCU mit Seriennummer, Firmware, Duty Cycle und Verbindungsstatus; das Standard-Interface steht kursiv an erster Stelle.

Darunter folgt je Kanal eine Zeile:

| Spalte | Inhalt |
| :----- | :----- |
| ## | laufende Nummer |
| IPS ID | Objekt-ID der Symcon-Instanz, „-“ bei Kanälen ohne Instanz |
| IPS device name | Name der Symcon-Instanz (mit `ShowLongIPSDeviceNames` samt Pfad) |
| HM address | Adresse des Kanals in der CCU. Ein vorangestelltes `*` kennzeichnet virtuelle Geräte der CCU, etwa Rauchmeldergruppen. |
| HM device name | Name des Kanals in der CCU (nur mit `ShowHMConfiguratorDeviceNames`) |
| HM device type | Gerätetyp, z. B. `HMIP-SWDO` |
| Fw. | Firmware des Geräts |
| HM channel type | Kanaltyp, z. B. `SHUTTER_CONTACT` |
| Dir. | `TX` = Kanal sendet (Sensor, Taster), `RX` = Kanal empfängt (Aktor) |
| AES | Verschlüsselung an (+) oder aus (-) |
| Roaming | Roaming an (+) oder aus (-) |
| je Interface | Pegelpaar in dBm |

Gerätetyp, Firmware, Roaming und Pegel stehen nur in der ersten Zeile eines Geräts; die weiteren Kanäle desselben Geräts lassen diese Spalten leer.

**Pegelpaare und Markierungen**

- Links steht der Pegel, mit dem das Gerät das Interface zuletzt empfangen hat, rechts der Pegel, mit dem das Interface das Gerät empfangen hat. Je näher an 0, desto besser. `--` heißt: kein Wert bekannt.
- **Unterstrichen:** das Interface, dem das Gerät zugeordnet ist; bei Roaming alle Interfaces.
- **Gelb:** das Interface mit dem besten linken Wert. Ist links bei keinem Interface ein Wert bekannt, trägt das erste Interface die Farbe – dann sagt sie nichts aus.
- **Rote Schrift:** Der Kanal ist mehr als einer Symcon-Instanz zugeordnet.
- Geräte ohne Pegel haben seit dem letzten Start des Funkdienstes der CCU nichts gesendet oder empfangen – oder sie gehören zu HmIP oder Wired, siehe [Grenzen](#7-grenzen).

## 4. Statusvariable

| Name | Ident | Darstellung | Bedeutung |
| :--- | :---- | :---------- | :-------- |
| Geräteliste | `DeviceList` | Wertanzeige | Die Zeilen des Reports als JSON, in derselben Reihenfolge. Nur mit `SaveDeviceListInVariable`; wird bei jedem Report neu geschrieben. |

Jeder Eintrag hat diese Felder:

| Schlüssel | Inhalt |
| :-------- | :----- |
| `IPS_occ` | laufende Nummer in der Reihenfolge der Erfassung |
| `IPS_id` | Objekt-ID der Symcon-Instanz, `"-"` bei Kanälen ohne Instanz |
| `IPS_name` | Name der Symcon-Instanz, `"-"` bei Kanälen ohne Instanz |
| `IPS_HM_d_assgnd` | `true`, wenn der Kanal mehr als einer Symcon-Instanz zugeordnet ist |
| `HM_address` | Adresse des Kanals |
| `HM_device` | Gerätetyp |
| `HM_devname` | Kanalname in der CCU (`"-"` ohne `ShowHMConfiguratorDeviceNames`) |
| `HM_FWversion` | Firmware |
| `HM_devtype` | Kanaltyp |
| `HM_direction` | `TX`, `RX` oder `-` |
| `HM_AES_active` | `+` oder `-` |
| `HM_Interface` | Seriennummer des zugeordneten Interfaces |
| `HM_Roaming` | `+` oder `-` |
| `HM_levels` | nur bei Geräten mit Pegeln: je verbundenem Interface, für das die CCU Werte hat, eine Liste `[Pegel am Gerät, Pegel am Interface, zugeordnet, gelb markiert]`; `65536` steht für „kein Wert“ |

## 5. Automatisieren

Für einen regelmäßigen Report genügt `UpdateInterval`. Wer die Daten weiterverarbeiten will, schaltet `SaveDeviceListInVariable` ein und liest die Variable. Ein Skript, das alle Funkgeräte meldet, die ihr zugeordnetes Interface schlechter als mit -85 dBm empfängt (der Grenzwert ist nur ein Beispiel):

```php
$hmi = 12345; // ID der Instanz HM Inventory Report Creator

if (!HMI_CreateReport($hmi)) {
    return;
}
$liste = json_decode(GetValue(IPS_GetObjectIDByIdent('DeviceList', $hmi)), true);
foreach ($liste as $eintrag) {
    foreach ($eintrag['HM_levels'] ?? [] as [$amGeraet, $amInterface, $zugeordnet, $gelb]) {
        if ($zugeordnet && $amInterface !== 65536 && $amInterface < -85) {
            IPS_LogMessage('HM Inventory', sprintf('%s (%s): %d dBm', $eintrag['HM_address'], $eintrag['HM_device'], $amInterface));
        }
    }
}
```

## 6. Funktionen für Skripte

```php
HMI_CreateReport(int $InstanzID): bool;
```
Erstellt den Report mit den Einstellungen der Instanz und schreibt, falls eingestellt, die Geräteliste. Gibt `false` zurück, wenn der Report nicht erstellt werden konnte (Meldungen siehe [Installation](#2-installation-und-erster-lauf)).

```php
HMI_GetOutputFileAbsolutePath(int $InstanzID): string;
```
Liefert den absoluten Pfad der Ausgabedatei; relative Angaben werden gegen das Symcon-Verzeichnis aufgelöst.

```php
HMI_GetReportUrl(int $InstanzID): string;
```
Liefert die Adresse, unter der der Report im Browser erreichbar ist, oder einen Leerstring, wenn die Datei nicht unterhalb von `user/` liegt.

## 7. Grenzen

- **Pegel gibt es nur für Homematic-Funkgeräte (BidCos).** Die CCU liefert sie nur für diesen Dienst. HmIP- und Wired-Geräte stehen im Report, aber ohne Pegel.
- **Ohne den BidCos-Funkdienst kein Report.** Die Liste der Interfaces kommt von diesem Dienst. Antwortet er nicht, bricht die Erstellung ab. Antwortet dagegen nur HmIP oder Wired nicht, fehlen deren Geräte, der Report entsteht trotzdem.
- **CCU-Namen kosten Zeit.** Für jeden Kanal geht eine eigene Anfrage an die CCU. Bei mehreren hundert Kanälen dauert der Report dadurch spürbar länger. Scheitert eine Anfrage, bleibt der Name leer und im Symcon-Log steht `CCU unreachable`.
- **Pegel sind Momentaufnahmen.** Die CCU merkt sich den letzten Pegel je Gerät seit dem Start ihres Funkdienstes. Ein Gerät, das seitdem geschwiegen hat, steht ohne Werte da.
- **Nur Instanzen am selben Socket.** Als „in Symcon angelegt“ zählen die HomeMatic-Geräteinstanzen, die an demselben HomeMatic Socket hängen wie diese Instanz.
- **Die Browser-Adresse kann danebenliegen.** Sie setzt sich aus der ersten IP-Adresse des Symcon-Rechners und dem Port der ersten aktiven WebServer-Instanz zusammen (ohne WebServer: Port 3777). Bei mehreren Netzwerkkarten oder einem VPN ist das womöglich nicht die Adresse, unter der der Rechner im Heimnetz erreichbar ist.
- **Der Report wird überschrieben.** Jeder Lauf ersetzt die Datei, eine Historie gibt es nicht.

## 8. Begriffe

| Begriff | Bedeutung |
| :------ | :-------- |
| BidCos | das Funkprotokoll der klassischen Homematic-Geräte (HM-RF); HmIP ist das neuere Protokoll von Homematic IP |
| Interface | Funkmodul, über das die CCU mit den Geräten spricht: das eingebaute Funkmodul oder ein LAN-Gateway |
| Kanal | Teilfunktion eines Geräts mit eigener Adresse (`<Seriennummer>:<Nummer>`); ein Zweifach-Aktor hat z. B. zwei Schaltkanäle |
| Duty Cycle | wie viel des erlaubten Sendezeit-Budgets (1 % je Stunde) ein Interface verbraucht hat; bei 100 % darf es vorübergehend nicht mehr senden |
| Roaming | Das Gerät ist keinem festen Interface zugeordnet, die CCU wählt das jeweils beste. |
| dBm | Einheit des Empfangspegels; Werte sind negativ, je näher an 0, desto stärker das Signal |

## 9. Anhang: Technik

- Die Daten kommen per XML-RPC von den Diensten der CCU: `listDevices` je Dienst (BidCos-RF, HmIP, BidCos-Wired), `listBidcosInterfaces` und `rssiInfo` vom BidCos-RF-Dienst. Die Kanalnamen holt ein HM-Script über den Skript-Port der CCU (`Script.exe`).
- Ports, SSL und Zugangsdaten stammen aus dem HomeMatic Socket. Mit SSL wird das selbstsignierte Zertifikat der CCU nicht geprüft.
- Verwendete Bibliothek: [phpxmlrpc](https://github.com/gggeek/phpxmlrpc) 4.11.5 (nur `src/`, eingebunden über den mitgelieferten Autoloader).

**GUIDs**

| | GUID |
| :--- | :--- |
| Bibliothek *HM Inventory Module* | `{240F4263-D2CB-49BC-AC00-3A9DC2CF3C10}` |
| Modul *HM Inventory Report Creator* | `{E3BEF9D8-23D4-47A8-B823-53BD7AF65CC3}` |
