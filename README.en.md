# HM Inventory

[![Checks](https://github.com/bumaas/HMInventory/actions/workflows/check.yml/badge.svg)](https://github.com/bumaas/HMInventory/actions/workflows/check.yml)

Creates an HTML overview of all Homematic devices of a CCU: every channel with address, device type, firmware and the matching Symcon instance, plus the radio levels between the devices and each radio interface of the CCU. Optionally, the same list is stored as JSON in a variable.

[Deutsche Version](README.md)

### Contents

1. [When do I need this?](#1-when-do-i-need-this)
2. [Installation and first run](#2-installation-and-first-run)
3. [Reading the report](#3-reading-the-report)
4. [Status variable](#4-status-variable)
5. [Automation](#5-automation)
6. [Script functions](#6-script-functions)
7. [Limitations](#7-limitations)
8. [Terms](#8-terms)
9. [Appendix: technical details](#9-appendix-technical-details)

## 1. When do I need this?

- **A device responds unreliably.** For every Homematic radio device (BidCos) the report shows the level to each interface of the CCU and marks the assigned one. Weak links stand out at a glance, as does a device attached to an unfavourable interface.
- **What is paired but not set up in Symcon?** Channels that exist in the CCU but have no Symcon instance show "-" instead of an object ID.
- **Finding duplicate instances.** If two Symcon instances are attached to the same Homematic channel, both are shown in red.
- **Inventory.** Which devices with which firmware exist, and what are they called in the CCU and in Symcon?

## 2. Installation and first run

**Requirements**

- Symcon 9.0 or later
- A CCU (CCU2, CCU3, RaspberryMatic or similar) connected to Symcon through a **HomeMatic Socket**. The module uses the socket's address, ports, SSL setting and credentials and needs none of its own.

**In Symcon**

1. Install the module from the Module Store: *Homematic Inventory Module*.
2. Create an instance *HM Inventory Report Creator* anywhere. It attaches to the HomeMatic Socket. With several CCUs, create one instance per CCU and assign it to the right socket via *Change gateway*.
3. Check the settings (caption in the form in brackets):

| Field | Default | Meaning |
| :---- | :------ | :------ |
| `active` (active) | on | An inactive instance does not create a report. |
| `OutputFile` (Output File) | `<Symcon directory>/user/HM_inventory.html` | Path and name of the HTML file. Relative paths start at the Symcon directory. Only files below `user/` can be opened in a browser. |
| `SortOrder` (Sort Order) | HM address | Row order: HM address, HM device type, HM channel type, HM device name or IPS device name |
| `ShowLongIPSDeviceNames` (Show long IPS Device Names) | off | Shows the name of the Symcon instance with its full path in the object tree. |
| `ShowHMConfiguratorDeviceNames` (Show HM Configurator Device Names) | on | Additional column with the channel name from the CCU. Takes time, see [Limitations](#7-limitations). |
| `ShowNotUsedChannels` (Show Channels which are not used) | on | Also lists channels that have no Symcon instance. |
| `ShowMaintenanceEntries` (Show Maintenance Channel) | on | Includes the maintenance channels (channel 0). Only applies while unused channels are shown. |
| `ShowVirtualKeyEntries` (Show Virtual Key Entries) | off | Includes the virtual keys of the CCU. Only applies while unused channels are shown. |
| `SaveDeviceListInVariable` (Save Device List in Variable) | off | Creates the variable *Device List* and writes the list into it as JSON with every report, see [Status variable](#4-status-variable). |
| `UpdateInterval` (Update Interval) | 0 | Recreates the report every n minutes. 0 = only on button press or by script. |

4. Apply and press **Create Report**. A progress bar runs; at the end the form reports "Report created:" with the path of the file. If the file is below `user/`, the button **Open report in browser** appears. It shows the address of the report.

**Instance status**

| Status | Meaning | What to do |
| :----- | :------ | :--------- |
| active | ready | – |
| inactive | The instance is switched off or the HomeMatic Socket is not active. | switch `active` on or check the socket |
| Output file is not creatable | The output file is empty, or its directory is missing or not writable. | correct the path or create the directory |

**Messages while creating the report**

| Message | Meaning |
| :------ | :------ |
| Gateway is not configured! | The instance is not attached to a HomeMatic Socket. |
| Instance is not active! | same as status *inactive* |
| File "…" not writable! | The file could not be written, e.g. because it is locked. |
| Error! | No report was created. The Symcon log states the reason, e.g. `Can't get any device information from the BidCos-Services`: the CCU did not answer on any service. |

## 3. Reading the report

The report starts with a **summary**: the number of radio interfaces (and how many are connected), the number of devices per service (HM-RF, HM-Wired, HmIP) and the number of Symcon instances attached to this HomeMatic Socket. Next to it are the **radio interfaces** of the CCU with serial number, firmware, duty cycle and connection state; the default interface is listed first, in italics.

Below follows one row per channel:

| Column | Content |
| :----- | :------ |
| ## | running number |
| IPS ID | object ID of the Symcon instance, "-" for channels without instance |
| IPS device name | name of the Symcon instance (with `ShowLongIPSDeviceNames` including its path) |
| HM address | address of the channel in the CCU. A leading `*` marks virtual devices of the CCU, such as smoke detector groups. |
| HM device name | name of the channel in the CCU (only with `ShowHMConfiguratorDeviceNames`) |
| HM device type | device type, e.g. `HMIP-SWDO` |
| Fw. | firmware of the device |
| HM channel type | channel type, e.g. `SHUTTER_CONTACT` |
| Dir. | `TX` = channel sends (sensor, button), `RX` = channel receives (actuator) |
| AES | encryption on (+) or off (-) |
| Roaming | roaming on (+) or off (-) |
| `<serial number> (dBm)` | one level pair per connected interface |

Device type, firmware, roaming and levels appear only in the first row of a device; further channels of the same device leave these columns empty.

**Level pairs and highlighting**

- The left value is the level at which the device last received the interface, the right value the level at which the interface received the device. The closer to 0, the better. `--` means no value is known.
- **Underlined:** the interface the device is assigned to; with roaming, all interfaces.
- **Yellow:** the interface that receives the device best (highest right value). Sensors, thermostats and valve drives in particular often do not report the left value, so it does not count. Switching actuators usually do – when checking an actuator, look at the left number next to it.
- **Red text:** the channel is assigned to more than one Symcon instance.
- Devices without levels have not sent or received anything since the CCU's radio service started – or they are HmIP or wired devices, see [Limitations](#7-limitations).

## 4. Status variable

| Name | Ident | Presentation | Meaning |
| :--- | :---- | :----------- | :------ |
| Device List | `DeviceList` | value presentation | The report rows as JSON, in the same order. Only with `SaveDeviceListInVariable`; rewritten with every report. |

Each entry has these fields:

| Key | Content |
| :-- | :------ |
| `IPS_occ` | running number in order of collection |
| `IPS_id` | object ID of the Symcon instance, `"-"` for channels without instance |
| `IPS_name` | name of the Symcon instance, `"-"` for channels without instance |
| `IPS_HM_d_assgnd` | `true` if the channel is assigned to more than one Symcon instance |
| `HM_address` | address of the channel |
| `HM_device` | device type |
| `HM_devname` | channel name in the CCU (`"-"` without `ShowHMConfiguratorDeviceNames`) |
| `HM_FWversion` | firmware |
| `HM_devtype` | channel type |
| `HM_direction` | `TX`, `RX` or `-` |
| `HM_AES_active` | `+` or `-` |
| `HM_Interface` | serial number of the assigned interface |
| `HM_Roaming` | `+` or `-` |
| `HM_levels` | only for devices with levels: per connected interface the CCU has values for, a list `[level at device, level at interface, assigned, yellow]`; `65536` means "no value" |

## 5. Automation

For a regular report, `UpdateInterval` is enough. To process the data, switch on `SaveDeviceListInVariable` and read the variable. A script that reports all radio devices received by their assigned interface below -85 dBm (the threshold is just an example):

```php
$hmi = 12345; // ID of the HM Inventory Report Creator instance

if (!HMI_CreateReport($hmi)) {
    return;
}
$list = json_decode(GetValue(IPS_GetObjectIDByIdent('DeviceList', $hmi)), true);
foreach ($list as $entry) {
    foreach ($entry['HM_levels'] ?? [] as [$atDevice, $atInterface, $assigned, $yellow]) {
        if ($assigned && $atInterface !== 65536 && $atInterface < -85) {
            IPS_LogMessage('HM Inventory', sprintf('%s (%s): %d dBm', $entry['HM_address'], $entry['HM_device'], $atInterface));
        }
    }
}
```

## 6. Script functions

```php
HMI_CreateReport(int $InstanceID): bool;
```
Creates the report with the instance settings and, if configured, writes the device list. Returns `false` if the report could not be created (messages see [Installation](#2-installation-and-first-run)).

```php
HMI_GetOutputFileAbsolutePath(int $InstanceID): string;
```
Returns the absolute path of the output file; relative paths are resolved against the Symcon directory.

```php
HMI_GetReportUrl(int $InstanceID): string;
```
Returns the address at which the report can be opened in a browser, or an empty string if the file is not below `user/`.

## 7. Limitations

- **Levels exist only for Homematic radio devices (BidCos).** The CCU provides them for this service only. HmIP and wired devices appear in the report, but without levels.
- **No report without the BidCos radio service.** The interface list comes from this service; if it does not answer, the report is aborted. If only HmIP or Wired does not answer, their devices are missing, but the report is still created.
- **CCU names take time.** Each channel needs its own request to the CCU. With several hundred channels, the report takes noticeably longer. If a request fails, the name stays empty and the Symcon log shows `CCU unreachable`.
- **Levels are snapshots.** The CCU keeps the last level per device since its radio service started. A device that has been silent since then shows no values.
- **Only instances on the same socket.** "Set up in Symcon" means HomeMatic device instances attached to the same HomeMatic Socket as this instance.
- **The browser address may be off.** It is built from the first IP address of the Symcon computer and the port of the first active WebServer instance (without WebServer: port 3777). With several network adapters or a VPN, this may not be the address under which the computer is reachable in the home network.
- **The report is overwritten.** Every run replaces the file; there is no history.

## 8. Terms

| Term | Meaning |
| :--- | :------ |
| BidCos | the radio protocol of classic Homematic devices (HM-RF); HmIP is the newer Homematic IP protocol |
| Interface | radio module through which the CCU talks to the devices: the built-in radio module or a LAN gateway |
| Channel | sub-function of a device with its own address (`<serial number>:<number>`); a two-way actuator, for example, has two switching channels |
| Duty cycle | how much of the permitted transmission budget (1 % per hour) an interface has used; at 100 % it may temporarily not transmit |
| Roaming | the device is not bound to a fixed interface; the CCU picks the best one |
| dBm | unit of the received level; values are negative, the closer to 0, the stronger the signal |

## 9. Appendix: technical details

- The data comes from the CCU services via XML-RPC: `listDevices` per service (BidCos-RF, HmIP, BidCos-Wired), `listBidcosInterfaces` and `rssiInfo` from the BidCos-RF service. The channel names are fetched by an HM script through the script port of the CCU (`Script.exe`).
- Ports, SSL and credentials are taken from the HomeMatic Socket. With SSL, the CCU's self-signed certificate is not verified.
- Library used: [phpxmlrpc](https://github.com/gggeek/phpxmlrpc) 4.11.5 (only `src/`, loaded through its own autoloader).

**GUIDs**

| | GUID |
| :--- | :--- |
| Library *HM Inventory Module* | `{240F4263-D2CB-49BC-AC00-3A9DC2CF3C10}` |
| Module *HM Inventory Report Creator* | `{E3BEF9D8-23D4-47A8-B823-53BD7AF65CC3}` |
