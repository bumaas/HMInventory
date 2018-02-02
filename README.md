# HMInventory

Modul für Symcon ab Version 5.0.

Erstellt einen Report mit allen Homematic Geräten.

## Dokumentation

**Inhaltsverzeichnis**

1. [Funktionsumfang](#1-funktionsumfang)  
2. [Voraussetzungen](#2-voraussetzungen)  
3. [Installation](#3-installation)  
4. [Funktionsreferenz](#4-funktionsreferenz)
5. [Konfiguration](#5-konfiguartion)  
6. [Anhang](#6-anhang)  

## 1. Funktionsumfang

Das Modul holt sich über die XML-RPC Schnittstellen die Daten der HM-Geräte ('listDevices') und ergänzt sie um die RSSI Werte ('rssiInfo' bzw. 'getParamset'). Die sortierten Daten werden dann als html-Datei aufbereitet.

Zusätzlich wird die Liste der Devices bei Bedarf noch json-encoded in eine Stringvariable geschrieben, falls man die Daten noch anderweitig verwenden möchte.  

### ausgewiesene Daten:  

* ++: laufende Nummer
* IPS ID: Objekt ID in Symcon
* HM address: Homematic Geräte Adresse
* HM device type: Homematic Typbezeichnung
* Fw.: Firmware Version
* HM channel type: Kanaltyp 
* Dir.: Sender (TX) oder Empfänger (RX)
* AES: Verschlüsselung ein-/ausgeschaltet (+/-)
* Roaming: Roaming ein-/ausgeschaltet (+/-)
* tx/rx (dbµV): Sende-/Empfangsstärke je Interface


## 2. Voraussetzungen

 - Symcon 5.0

## 3. Installation

### a. Laden des Moduls

In Symcon unter Kerninstanzen über `Modules -> Hinzufügen` das Modul ergänzen mit der URL:
	
    `https://github.com/Wolbolar/IPSymconAstronomy`  

### b. Anlegen einer Instanz

In Symcon an beliebiger Stelle `Instanz hinzufügen` auswählen und `HM Inventory Report Creator` auswählen.
	

## 4. Funktionsreferenz


## 5. Konfiguration

### HM Inventory Report Creator

| Eigenschaft | Typ     | Standardwert            | Funktion                                  |
| :--------- | :-----: | :------------------------| :--------------------------------------- |
| Host        | string  | IP-Adresse des ersten gefundenen HomeMatic Socket | IP Adresse des Homematic Interfaces       |                  |
| OutputFile  | string  | 'HM_inventory.html' im IPS Kernel Verzeichnis     | wenn ein Dateiname angegeben ist, wird die Ausgabe im HTML Format in diese Datei geschrieben |
| SortOrder   | integer | 0                       | Sortierreihenfolge der HM Devices in der Ausgabeliste:<br>0 - HM address (default)<br>1 - HM device type<br>2 - HM channel type<br>3 - IPS device name |
| ShowLongIPSDeviceNames   | boolean | false      | legt fest, ob IPS Namen mit oder ohne vollständigem Pfad ausgegeben werden sollen |
| ShowVirtualKeyEntries   | boolean | false       | legt fest, ob die Virtuellen Kanäle der Homematic ausgegeben werden sollen |
| ShowHMConfiguratorDeviceNames | boolean | true  | legt fest, ob die in der Homematic gewählten Bezeichnungen ausgegeben werden sollen |
| ShowMaintenanceEntries | boolean | true         | legt fest, ob die MAINTENANCE (0) Kanäle ausgegeben werden sollen |
| SaveDeviceListInVariable | boolean | false      | legt fest, ob die Liste der gefundenen Devices json codiert in einer Stringvariablen gespeichert werden soll|
| UpdateIntervall | integer | 0                   | legt fest, in welchem regelmäßigen Abstand (in Minuten) der Report aufbereitet werden soll (0: deaktiviert)|


## 6. Anhang

###  a. Funktionen:

```php
HMI_CreateReport(int $InstanceID)
```
Erstellt den Report mit allen Homeatic Devices entsprechend der in der Instanz eingestellten Eigenschaften.



###  b. GUIDs und Datenaustausch

#### HM Inventory Modul

GUID: `{240F4263-D2CB-49BC-AC00-3A9DC2CF3C10}` 

#### HM Inventory Report Creator

GUID: `{E3BEF9D8-23D4-47A8-B823-53BD7AF65CC3}` 



