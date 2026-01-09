<?php /** @noinspection CurlSslServerSpoofingInspection */

declare(strict_types=1);

// We need the "xmlrpc" include file
// see https://github.com/gggeek/phpxmlrpc/releases
include_once __DIR__ . '/../libs/phpxmlrpc-4.3.0/lib/xmlrpc.inc';

// Klassendefinition

/** @noinspection AutoloadingIssuesInspection */
class HMInventoryReportCreator extends IPSModuleStrict
{
    private const array  SERVICETYPES = ['RF', 'IP', 'WR'];
    private const string ERROR_MSG    = "Can't get any device information from the BidCoS-%s-Service";

    // Some color options for the HTML output
    private const string BG_COLOR_INTERFACE_LIST = '#223344';         // Background color for the interface list
    private const int    INVALID_LEVEL           = 65536;

    //property names
    private const string PROP_ACTIVE                   = 'active';
    private const string PROP_OUTPUTFILE               = 'OutputFile';
    private const string PROP_SHOWMAINTENANCEENTRIES   = 'ShowMaintenanceEntries';
    private const string PROP_SHOWVIRTUALKEYENTRIES    = 'ShowVirtualKeyEntries';
    private const string PROP_SAVEDEVICELISTINVARIABLE = 'SaveDeviceListInVariable';
    private const string PROP_SORTORDER                 = 'SortOrder';
    private const string PROP_SHOWLONGIPSDEVICENAMES    = 'ShowLongIPSDeviceNames';
    private const string PROP_SHOWHMCONFIGURATORDEVICENAMES = 'ShowHMConfiguratorDeviceNames';
    private const string PROP_SHOWNOTUSEDCHANNELS       = 'ShowNotUsedChannels';
    private const string PROP_UPDATEINTERVAL            = 'UpdateInterval';


    // Überschreibt die interne IPS_Create($id) Funktion
    public function Create(): void
    {
        // Diese Zeile nicht löschen.
        parent::Create();

        $this->RegisterProperties();

        $this->RegisterTimer('Update', 0, 'HMI_CreateReport(' . $this->InstanceID . ');');

        //we will wait until the kernel is ready
        $this->RegisterMessage(0, IPS_KERNELMESSAGE);
    }

    public function ApplyChanges(): void
    {
        //Never delete this line!
        parent::ApplyChanges();

        //Set a receive filter to something that will never match
        $this->SetReceiveDataFilter(
            'Dieser Filter sollte niemals greifen, daher ist er etwas länger :-'
        ); //es werden keine Nachrichten vom verbundenen Socket verarbeitet

        if (IPS_GetKernelRunlevel() !== KR_READY) {
            return;
        }

        $this->SetTimerInterval('Update', $this->ReadPropertyInteger(self::PROP_UPDATEINTERVAL) * 60 * 1000);

        $this->RegisterVariables();

        $this->SetInstanceStatus();

        $this->SetSummary($this->ReadPropertyString(self::PROP_OUTPUTFILE));
    }

    public function MessageSink(int $TimeStamp, int $SenderID, int $Message, array $Data): void
    {
        parent::MessageSink($TimeStamp, $SenderID, $Message, $Data);

        if (($Message === IPS_KERNELMESSAGE) && ($Data[0] === KR_READY)) {
            $this->ApplyChanges();
        }
    }

    public function ReceiveData(string $JSONString): string
    {
        trigger_error(sprintf('Fatal error: no ReceiveData expected. (%s)', $JSONString));

        return '';
    }


    /**
     * Erstellt den HM Inventory Report und speichert ihn als HTML-Datei.
     *
     * Ursprünglich entwickelt von Andreas Bahrdt (HM-Inventory), Public Domain.
     * Quellen: http://www.ip-symcon.de/forum/99682-post76.html
     *          https://www.symcon.de/forum/threads/17633-HM_Inventory
     *
     * Historie:
     * - 27.10.2011: Anpassung für IPS v2.5 (Raketenschnecke)
     * - 16.03.2016: Anpassung für IPS v4.0 (bumaas)
     * - 18.01.2017: Erweiterung für HM-IP und HM-Wired (bumaas)
     *
     * @return bool True bei Erfolg, false falls der Report nicht erstellt werden konnte.
     * @throws \JsonException
     * @throws \JsonException
     */
    public function CreateReport(): bool
    {

        $this->progressBarInit();
        $progressBarCounter = 0;

        $this->UpdateFormField('ProgressBar', 'current', $progressBarCounter++); // Schritt 1

        $reportData = $this->collectDeviceData($progressBarCounter);
        if ($reportData === null) {
            return false;
        }

        // Wenn gewünscht, die sortierte Liste in die Variable schreiben
        if ($this->ReadPropertyBoolean(self::PROP_SAVEDEVICELISTINVARIABLE)) {
            $this->SetValue('DeviceList', json_encode($reportData['HM_array'], JSON_THROW_ON_ERROR));
        }

        // Generate HTML output code
        $this->UpdateFormField('ProgressBar', 'current', $progressBarCounter++); // Schritt 6

        // Header-Daten vorbereiten
        $headerHtml = $this->renderHeaderSection($reportData);

        // Interface-Liste
        $interfacesHtml = $this->renderInterfacesSection($reportData['hm_BidCos_Ifc_list']);

        // Geräte-Liste
        $devicesHtml = $this->renderDevicesSection($reportData);

        $this->UpdateFormField('ProgressBar', 'current', $progressBarCounter++); // Schritt 7

        $notesHtml = $this->renderNotesSection();



        // Output the results
        $outputFileName = $this->ReadPropertyString(self::PROP_OUTPUTFILE);
        if ($outputFileName) {
            $this->UpdateFormField('ProgressBar', 'current', $progressBarCounter++); // Schritt 8

            $htmlContent = $this->getHtmlContent(
                $headerHtml,
                $interfacesHtml,
                $this->renderSeparator(),
                $devicesHtml,
                $notesHtml,
                '</table>'
            );

            if (@file_put_contents($outputFileName, $htmlContent) === false) {
                echo sprintf('File "%s" not writable!' . PHP_EOL, $outputFileName);
                return false;
            }

            $this->UpdateFormField('ProgressBar', 'current', 9); // Finaler Schritt 9
            IPS_Sleep(200);
            $this->UpdateFormField('ProgressBar', 'visible', false);
            return true;
        }

        return false;
    }

    /**
     * Sammelt alle benötigten Daten für den Report.
     *
     * @param int &$progressBarCounter Referenz auf den globalen Fortschrittszähler
     *
     * @return array|null
     * @throws \JsonException
     * @throws \JsonException
     */
    private function collectDeviceData(int &$progressBarCounter): ?array
    {
        $ParentId = $this->fetchParentId();
        if (!$this->isGatewayActive($ParentId)) {
            return null;
        }

        $ParentConfig = json_decode(IPS_GetConfiguration($ParentId), true, 512, JSON_THROW_ON_ERROR);
        $IP_adr_Homematic = IPS_GetProperty($ParentId, 'Host');
        [$BidCos_RF_Service_adr, $BidCos_IP_Service_adr] = $this->formatServiceAddresses($ParentConfig, $IP_adr_Homematic);

        [$hm_dev_list, $dev_counter, $err] = $this->getDeviceLists($ParentId, $ParentConfig);
        if (count($hm_dev_list) === 0) {
            trigger_error("Can't get any device information from the BidCos-Services (Error: $err)");
            return null;
        }

        // 1. BidCos Interfaces
        $this->UpdateFormField('ProgressBar', 'current', $progressBarCounter++);  // Schritt 1
        $xml_rtnmsg = $this->SendRequestMessage('listBidcosInterfaces', [], $BidCos_RF_Service_adr, $ParentConfig['UseSSL'], $ParentConfig['Password'], $ParentConfig['Username']);

        if ($xml_rtnmsg->errno !== 0) {
            $this->SendDebug('Error', "Can't get HM-interface information from the BidCos-RF-Service", 0);
            return null;
        }

        $hm_BidCos_Ifc_list = php_xmlrpc_decode($xml_rtnmsg->value());
        $default = array_column($hm_BidCos_Ifc_list, 'DEFAULT');
        array_multisort($default, SORT_DESC, $hm_BidCos_Ifc_list);

        $HM_interface_num = 0;
        $HM_interface_connected_num = 0;
        $HM_default_interface_no = 0;
        foreach ($hm_BidCos_Ifc_list as $key => $hm_ifce) {
            $HM_interface_num++;
            if ($hm_ifce['CONNECTED']) {
                $HM_interface_connected_num++;
            }
            if ($hm_ifce['DEFAULT']) {
                $HM_default_interface_no = $key;
            }
        }

        $IPS_device_num = 0;
        $IPS_HM_channel_num = 0;
        $HM_module_num = 0;
        $HM_array = [];

        // 2. Fill HM_array with devices found in IP-Symcon
        $this->UpdateFormField('ProgressBar', 'current', $progressBarCounter++);  // Schritt 2
        foreach (IPS_GetInstanceListByModuleID('{EE4A81C6-5C90-4DB7-AD2F-F6BBD521412E}') as $id) {
            if ($ParentId !== IPS_GetInstance($id)['ConnectionID']) {
                continue;
            }
            $HM_module_num++;
            $IPS_device_num++;
            $IPS_HM_channel_already_assigned = false;
            $HM_address = IPS_GetProperty($id, 'Address');

            $NeedlePos = strpos($HM_address, ':');
            if (!$NeedlePos) {
                continue;
            }
            $HM_Par_address = substr($HM_address, 0, $NeedlePos);

            $HM_device = '-'; $HM_devname = '-'; $HM_FWversion = ' '; $HM_Interface = '';
            $HM_Roaming = ' '; $HM_devtype = '-'; $HM_direction = '-'; $HM_AES_active = '-';
            $hm_chld_dev = null; $hm_par_dev = null;

            foreach ($hm_dev_list as $hm_dev) {
                match ($hm_dev['ADDRESS']) {
                    $HM_address     => $hm_chld_dev = $hm_dev,
                    $HM_Par_address => $hm_par_dev = $hm_dev,
                    default         => null
                };

                if ($hm_chld_dev !== null) {
                    if (isset($hm_dev['PARENT_TYPE'])) $HM_device = $hm_dev['PARENT_TYPE'];
                    if ($this->ReadPropertyBoolean(self::PROP_SHOWHMCONFIGURATORDEVICENAMES)) {
                        $HM_devname = $this->getHMChannelName($IP_adr_Homematic, $hm_dev['ADDRESS']);
                    }
                    $device_info = $this->extractDeviceInfo($hm_par_dev);
                    $HM_FWversion = $device_info['FIRMWARE'];
                    $HM_Interface = $device_info['INTERFACE'];
                    $HM_Roaming   = $device_info['ROAMING'];
                    $HM_devtype = $hm_dev['TYPE'];
                    if (isset($hm_dev['DIRECTION'])) {
                        if ($hm_dev['DIRECTION'] === 1) $HM_direction = 'TX';
                        elseif ($hm_dev['DIRECTION'] === 2) $HM_direction = 'RX';
                    }
                    if (isset($hm_dev['AES_ACTIVE']) && ($hm_dev['AES_ACTIVE'] !== 0)) $HM_AES_active = '+';
                    break;
                }
            }

            if ($HM_address !== '') {
                foreach ($HM_array as &$HM_dev_ref) {
                    if ($HM_dev_ref['HM_address'] === $HM_address) {
                        $HM_dev_ref['IPS_HM_d_assgnd'] = true;
                        $IPS_HM_channel_already_assigned = true;
                        break;
                    }
                }
                unset($HM_dev_ref);
                if (!$IPS_HM_channel_already_assigned) $IPS_HM_channel_num++;
            }

            $IPS_name = $this->ReadPropertyBoolean(self::PROP_SHOWLONGIPSDEVICENAMES) ? IPS_GetLocation($id) : IPS_GetName($id);

            $HM_array[] = [
                'IPS_occ' => $HM_module_num, 'IPS_id' => $id, 'IPS_name' => $IPS_name,
                'IPS_HM_d_assgnd' => $IPS_HM_channel_already_assigned, 'HM_address' => $HM_address,
                'HM_device' => $HM_device, 'HM_devname' => $HM_devname, 'HM_FWversion' => $HM_FWversion,
                'HM_devtype' => $HM_devtype, 'HM_direction' => $HM_direction, 'HM_AES_active' => $HM_AES_active,
                'HM_Interface' => $HM_Interface, 'HM_Roaming' => $HM_Roaming
            ];
        }

        // 3. Add devices not in Symcon (optional)
        $this->UpdateFormField('ProgressBar', 'current', $progressBarCounter++);  // Schritt 3
        if ($this->ReadPropertyBoolean(self::PROP_SHOWNOTUSEDCHANNELS)) {
            foreach ($hm_dev_list as $hm_dev) {
                $HM_address      = $hm_dev['ADDRESS'];
                $hm_dev_in_array = false;
                foreach ($HM_array as $HM_dev_a) {
                    if ($hm_dev['ADDRESS'] === $HM_dev_a['HM_address']) {
                        $hm_dev_in_array = true;
                        break;
                    }
                }

                if (($hm_dev_in_array === false) && isset($hm_dev['PARENT']) && ($hm_dev['PARENT'] !== '')) {
                    if ($hm_dev['TYPE'] === 'VIRTUAL_KEY' && !$this->ReadPropertyBoolean(self::PROP_SHOWVIRTUALKEYENTRIES)) {
                        continue;
                    }
                    if ($hm_dev['TYPE'] === 'MAINTENANCE' && !$this->ReadPropertyBoolean(self::PROP_SHOWMAINTENANCEENTRIES)) {
                        continue;
                    }
                    $hm_chld_dev    = $hm_dev;
                    $hm_par_dev     = null;
                    $HM_Par_address = substr($HM_address, 0, strpos($HM_address, ':'));
                    foreach ($hm_dev_list as $hm_p_dev) {
                        if ($hm_p_dev['ADDRESS'] === $HM_Par_address) {
                            $hm_par_dev = $hm_p_dev;
                        }
                    }
                    if ($hm_chld_dev !== null) {
                        $HM_module_num++;
                        $HM_device  = $hm_chld_dev['PARENT_TYPE'];
                        $HM_devname = '-';
                        if ($this->ReadPropertyBoolean('ShowHMConfiguratorDeviceNames')) {
                            $HM_devname = $this->getHMChannelName($IP_adr_Homematic, $HM_address);
                        }

                        $device_info = $this->extractDeviceInfo($hm_par_dev);

                        $HM_FWversion = $device_info['FIRMWARE'];
                        $HM_Interface = $device_info['INTERFACE'];
                        $HM_Roaming   = $device_info['ROAMING'];

                        $HM_devtype    = $hm_chld_dev['TYPE'];
                        $HM_direction  = '-';
                        $HM_AES_active = '-';
                        if ($hm_dev['DIRECTION'] === 1) {
                            $HM_direction = 'TX';
                        } elseif ($hm_dev['DIRECTION'] === 2) {
                            $HM_direction = 'RX';
                        }
                        if ($hm_chld_dev['AES_ACTIVE'] !== 0) {
                            $HM_AES_active = '+';
                        }

                        $HM_array[] = [
                            'IPS_occ'         => $HM_module_num,
                            'IPS_id'          => '-',
                            'IPS_name'        => '-',
                            'IPS_HM_d_assgnd' => false,
                            'HM_address'      => $HM_address,
                            'HM_device'       => $HM_device,
                            'HM_devname'      => $HM_devname,
                            'HM_FWversion'    => $HM_FWversion,
                            'HM_devtype'      => $HM_devtype,
                            'HM_direction'    => $HM_direction,
                            'HM_AES_active'   => $HM_AES_active,
                            'HM_Interface'    => $HM_Interface,
                            'HM_Roaming'      => $HM_Roaming
                        ];
                    }
                }
            }
        }

        // 4. RSSI Levels (RF & IP)
        $this->UpdateFormField('ProgressBar', 'current', $progressBarCounter++);  // Schritt 4
        // Hole RSSI Infos vom RF Service
        $xml_rtnmsg = $this->SendRequestMessage('rssiInfo', [], $BidCos_RF_Service_adr, $ParentConfig['UseSSL'], $ParentConfig['Password'], $ParentConfig['Username']);
        $hm_rssi_list = ($xml_rtnmsg->errno === 0) ? php_xmlrpc_decode($xml_rtnmsg->value()) : [];

        foreach ($HM_array as &$HM_dev) {
            $hm_adr = explode(':', $HM_dev['HM_address']);

            // Nur für das Hauptgerät (Kanal 0 oder ohne Kanal) RSSI suchen
            if (isset($hm_rssi_list[$hm_adr[0]])) {
                $HM_dev['HM_levels'] = [];
                foreach ($hm_BidCos_Ifc_list as $ifc) {
                    if (!$ifc['CONNECTED']) continue;

                    $adr = $ifc['ADDRESS'];
                    if (isset($hm_rssi_list[$hm_adr[0]][$adr])) {
                        // Speichere RX, TX, IsAssociated, IsBest
                        $HM_dev['HM_levels'][] = [
                            $hm_rssi_list[$hm_adr[0]][$adr][0], // RX
                            $hm_rssi_list[$hm_adr[0]][$adr][1], // TX
                            ($HM_dev['HM_Interface'] === $adr), // IsAssociated
                            false // IsBest (wird unten berechnet)
                        ];
                    }
                }

                // Bestes Interface markieren
                if (count($HM_dev['HM_levels']) > 0) {
                    $best_idx = 0;
                    $max_rssi = -255;
                    foreach ($HM_dev['HM_levels'] as $idx => $lvl) {
                        if ($lvl[0] > $max_rssi && $lvl[0] !== self::INVALID_LEVEL) {
                            $max_rssi = $lvl[0];
                            $best_idx = $idx;
                        }
                    }
                    $HM_dev['HM_levels'][$best_idx][3] = true;
                }
            }
        }
        unset($HM_dev);

        // Sortierung anwenden
        $sortOrder = $this->ReadPropertyInteger(self::PROP_SORTORDER);
        switch ($sortOrder) {
            case 1: // HM device type
                usort($HM_array, [self::class, 'usort_HM_devtype']);
                break;
            case 2: // HM channel type
                usort($HM_array, [self::class, 'usort_HM_device_adr']);
                break;
            case 3: // IPS device name
                usort($HM_array, [self::class, 'usort_IPS_dev_name']);
                break;
            case 4: // HM device name
                usort($HM_array, [self::class, 'usort_HM_devname']);
                break;
            case 0: // HM address
            default:
                usort($HM_array, [self::class, 'usort_HM_address']);
                break;
        }

        return [
            'dev_counter' => $dev_counter,
            'HM_array' => $HM_array,
            'hm_BidCos_Ifc_list' => $hm_BidCos_Ifc_list,
            'HM_interface_num' => $HM_interface_num,
            'HM_interface_connected_num' => $HM_interface_connected_num,
            'IPS_device_num' => $IPS_device_num,
            'IPS_HM_channel_num' => $IPS_HM_channel_num,
            'HM_module_num' => $HM_module_num,
            'HM_default_interface_no' => $HM_default_interface_no
        ];
    }

    private function renderHeaderSection(array $data): string
    {
        $moduleVersion = $this->getModuleVersion();
        $inventoryStr = sprintf('<b>HM Inventory (%s) </b><b>&nbsp found at %s</b>', $moduleVersion, date('d.m.Y H:i:s'));
        $interfaceStr = sprintf(
            '%s HomeMatic interfaces (%s connected) with %s HM-RF devices, %s HM-wired devices and %s HmIP devices',
            $data['HM_interface_num'],
            $data['HM_interface_connected_num'],
            $data['dev_counter']['RF'],
            $data['dev_counter']['WR'],
            $data['dev_counter']['IP']
        );
        $instanceStr = sprintf('%s IPS instances (connected to %s HM channels)', $data['IPS_device_num'], $data['IPS_HM_channel_num']);

        $html = "<table class='table-align-left Background-color'>";
        $html .= "<tr style='vertical-align: top'><td><table style='text-align: left; font-size: large; color: #99AABB'>";
        $html .= $this->generateTableRow('large', '#99AABB', $inventoryStr);
        $html .= $this->generateTableRow('small', '#CCCCCC', $interfaceStr);
        $html .= $this->generateTableRow('small', '#CCCCCC', $instanceStr);
        $html .= '</table></td>';
        return $html;
    }

    private function renderInterfacesSection(array $interfaceList): string
    {
        $html = "<td style='vertical-align: top'>&nbsp;</td>";
        $html .= "<td style='width: 40%; vertical-align: bottom;'><table style='width: 100%; text-align: right; background-color: "
                 . self::BG_COLOR_INTERFACE_LIST . "'>";
        foreach ($interfaceList as $ifc) {
            $html .= $this->formatInterfaceRow($ifc);
        }
        $html .= '</table></td></tr>';
        return $html;
    }

    private function renderDevicesSection(array $data): string
    {
        $html = '<tr><td colspan=3><table class="table-align-left">';
        $html .= '<tr class="bgcolor-header-devices">';

        // Header-Spalten
        $dthdr_td_b = '<td style="font-size: small; color: #EEEEEE"><b>';
        $dthdr_td_b_r = '<td style="text-align: right; font-size: small; color: #EEEEEE"><b>';
        $dthdr_td_e = '</b></td>';
        $dthdr_td_eb = $dthdr_td_e . $dthdr_td_b;

        $html .= $dthdr_td_b_r . '&nbsp##&nbsp' . $dthdr_td_eb . 'IPS ID' . $dthdr_td_eb . 'IPS device name&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp'
                 . $dthdr_td_eb . 'HM address' . $dthdr_td_e;

        if ($this->ReadPropertyBoolean(self::PROP_SHOWHMCONFIGURATORDEVICENAMES)) {
            $html .= $dthdr_td_b . 'HM device name' . $dthdr_td_e;
        }

        $html .= $dthdr_td_b . 'HM device type' . $dthdr_td_eb . 'Fw.' . $dthdr_td_eb . 'HM channel type' . $dthdr_td_eb . 'Dir.' . $dthdr_td_eb . 'AES' . $dthdr_td_e;
        $html .= '<td style="width: 2%; text-align: center; color: #EEEEEE; font-size: medium">Roa- ming</td>';

        foreach ($data['hm_BidCos_Ifc_list'] as $hm_ifce) {
            if ($hm_ifce['CONNECTED']) {
                $html .= '<td style="width: 6%; text-align: center; color: #EEEEEE; font-size: small">' . $hm_ifce['ADDRESS'] . ' tx/rx&nbsp(db&micro;V)</td>';
            }
        }
        $html .= '</tr>' . PHP_EOL;

        $entry_no = 0;
        $previous_hm_adr = '';
        foreach ($data['HM_array'] as $HM_dev) {
            $hm_adr = explode(':', $HM_dev['HM_address']);
            $same_device = ($hm_adr[0] === $previous_hm_adr);
            $previous_hm_adr = $hm_adr[0];

            $html .= $this->renderDeviceRow($HM_dev, ++$entry_no, $same_device, $data['HM_interface_connected_num']);
        }

        if ($data['HM_module_num'] === 0) {
            $html .= '<tr><td colspan=20 style="text-align: center; color: #DDDDDD; font-size: large"><br/>No HomeMatic devices found!</td></tr>';
        }

        $html .= '</table></td></tr>';
        return $html;
    }

    private function renderDeviceRow(array $dev, int $entryNo, bool $sameDevice, int $connectedIfcNum): string
    {
        $td_color = ($dev['IPS_HM_d_assgnd'] === false) ? 'gray-text' : 'pale-red-text';
        $tr_class = (($entryNo % 2) === 1) ? 'bg_color_oddline' : 'bg_color_evenline';

        $html = '<tr class="' . $tr_class . '">';
        $html .= '<td class="' . $td_color . ' text-right">' . $entryNo . '&nbsp&nbsp</td>';
        $html .= '<td class="' . $td_color . '">' . $dev['IPS_id'] . '</td>';
        $html .= '<td class="' . $td_color . '">' . $dev['IPS_name'] . '</td>';
        $html .= '<td class="' . $td_color . '">' . $dev['HM_address'] . '</td>';

        if ($this->ReadPropertyBoolean(self::PROP_SHOWHMCONFIGURATORDEVICENAMES)) {
            $html .= '<td class="' . $td_color . '">' . $dev['HM_devname'] . '</td>';
        }

        if (!$sameDevice) {
            $html .= '<td class="' . $td_color . '">' . $dev['HM_device'] . '</td>';
            $html .= '<td class="' . $td_color . '">' . $dev['HM_FWversion'] . '</td>';
        } else {
            // Platzhalter für FW und Device-Typ bei Folgekanälen
            $html .= '<td class="' . $td_color . '"></td><td class="' . $td_color . '"></td>';
        }

        $html .= '<td class="' . $td_color . '">' . $dev['HM_devtype'] . '</td>';
        $html .= '<td class="' . $td_color . '">' . $dev['HM_direction'] . '</td>';
        $html .= '<td class="' . $td_color . ' text-center">' . $dev['HM_AES_active'] . '</td>';

        if (!$sameDevice) {
            $html .= '<td class="' . $td_color . ' text-center">' . $dev['HM_Roaming'] . '</td>';
            $html .= $this->renderRssiColumns($dev, $connectedIfcNum);
        } else {
            // WICHTIG: Auch hier müssen leere Zellen für Roaming + alle Interfaces gerendert werden
            $html .= '<td class="' . $td_color . '"></td>';
            $html .= str_repeat('<td></td>', $connectedIfcNum);
        }

        $html .= '</tr>' . PHP_EOL;
        return $html;
    }

    private function renderNotesSection(): string
    {
        return <<<HEREDOC
        <tr><!-- Notes Überschriftenzeile-->
            <td colspan="3">
                <table style="width: 100%; text-align: left; font-size:medium; color: #DDDDDD">
                    <tr>
                        <td>Notes:</td>
                    </tr>
                    <tr>
                        <td style="font-size: smaller; color: #DDDDDD">
                            <ol>
                                <li>Interfaces: bold letters indicate the default BidCos-Interface.</li>
                                <li>Level-pairs: the left value is showing the last signal level received by the device from the interface, while the
                                    right value is showing the last signal level received by the interface from the device.
                                </li>
                                <li>Level-pairs: underlined letters of the level-pair indicate the BidCos-Interface associated with the device (or all
                                    interfaces when Roaming is enabled for the device).
                                </li>
                                <li>Level-pairs: the yellow level-pair indicates the BidCos-Interface with best signal quality.</li>
                                <li>Devices without level-pairs haven't sent/received anything since last start of the BidCos-service or are wired.</li>
                                <li>BidCos channels assigned to more than one IPS-device are shown in red.</li>
                            </ol>
                        </td>
                    </tr>
                </table>
            </td>
        </tr><!-- Ende Notes Überschriftenzeile-->
HEREDOC;
    }

    private function renderRssiColumns(array $dev, int $connectedIfcNum): string
    {
        $html = '';
        if (!isset($dev['HM_levels'])) {
            return str_repeat('<td></td>', $connectedIfcNum);
        }

        for ($i = 0; $i < $connectedIfcNum; $i++) {
            if (isset($dev['HM_levels'][$i])) {
                $lciValue = $dev['HM_levels'][$i];
                $isBest = (bool)($lciValue[3] ?? false);
                $isAssociated = ($dev['HM_Roaming'] === '+') || ($lciValue[2] ?? false);

                if ($isBest) {
                    $color = $isAssociated ? '#DDDD66' : '#FFFF88';
                } else {
                    $color = '#DDDDDD';
                }

                [$rx, $tx] = $this->getRxTxLevelString((int)$lciValue[0], (int)$lciValue[1]);

                $content = $isAssociated ? "<ins>$rx &#047 $tx</ins>" : "$rx &#047 $tx";
                $html .= '<td class="text-center"><p style="color: ' . $color . '">' . $content . '</p></td>';
            } else {
                $html .= '<td></td>';
            }
        }
        return $html;
    }

    private function renderSeparator(): string
    {
        return '<tr><td colspan=3><table class="table-align-left"><tr><td></td></tr></table></td></tr>';
    }

    private function extractDeviceInfo($hm_par_dev): array
    {
        if ($hm_par_dev !== null) {
            $HM_FWversion = $hm_par_dev['FIRMWARE'];
            $HM_Interface = $hm_par_dev['INTERFACE'] ?? '';
            if (isset($hm_par_dev['ROAMING']) && $hm_par_dev['ROAMING']) {
                $HM_Roaming = '+';
            } else {
                $HM_Roaming = '-';
            }
        } else {
            $HM_FWversion = '';
            $HM_Interface = '';
            $HM_Roaming   = '';
        }
        return ['FIRMWARE' => $HM_FWversion, 'INTERFACE' => $HM_Interface, 'ROAMING' => $HM_Roaming];
    }


    private function generateTableRow($fontSize, $color, $content): string
    {
        return "<tr><td style='font-size: $fontSize; color: $color'>$content</td>";
    }

    private function formatInterfaceRow($hm_ifce): string
    {
        $italicStart = $hm_ifce['DEFAULT'] ? '<i>' : '';
        $italicEnd   = $hm_ifce['DEFAULT'] ? '</i>' : '';

        $connectionStatus = $hm_ifce['CONNECTED'] ? 'connected' : 'Not connected';
        $interfaceInfo    = sprintf('%s (Fw: %s, DC: %s%%)', $hm_ifce['ADDRESS'], $hm_ifce['FIRMWARE_VERSION'], $hm_ifce['DUTY_CYCLE']);

        return "<tr>
                <td style='font-size: small;color: #EEEEEE'>{$italicStart}Interface: $interfaceInfo&nbsp{$italicEnd}</td>
                <td style='font-size: small;color: #EEEEEE'>{$italicStart}{$hm_ifce['DESCRIPTION']}{$italicEnd}</td>
                <td style='font-size: small;color: #EEEEEE'>{$italicStart}{$connectionStatus}{$italicEnd}</td>
            </tr>" . PHP_EOL;
    }

    private function getRxTxLevelString(int $rx_lvl, int $tx_lvl): array
    {
        return [
            $rx_lvl !== self::INVALID_LEVEL ? (string)$rx_lvl : '--',
            $tx_lvl !== self::INVALID_LEVEL ? (string)$tx_lvl : '--'
        ];
    }

    /**
     * Generates the HTML content for the report.
     *
     * @param string $HTML_intro The HTML content for the introduction section.
     * @param string $HTML_ifcs  The HTML content for the interfaces section.
     * @param string $HTML_sep   The HTML content for the separators section.
     * @param string $HTML_dvcs  The HTML content for the devices section.
     * @param string $HTML_notes The HTML content for the notes section.
     * @param string $HTML_end   The HTML content for the ending section.
     *
     * @return string The generated HTML content.
     */
    private function getHtmlContent(
        string $HTML_intro,
        string $HTML_ifcs,
        string $HTML_sep,
        string $HTML_dvcs,
        string $HTML_notes,
        string $HTML_end
    ): string {
        return <<<HEREDOC
<html lang="">
<head>
    <style>
        html, body {
            font-family: Arial, Helvetica, sans-serif;
            font-size: 12px;
            background-color: #000000;
            color: #dddddd;
        }
        
        .table-align-left {
            width: 100%;
            text-align: left;
        }
        
        .Background-color {
            background-color: #223344;
        }
        
        .bgcolor-header-devices {
            background-color: #334455;
        }
        
        .bg_color_oddline {
            background-color: #181818;
            font-size: 0.8em;
        }

        .bg_color_evenline {
            background-color: #1A2B3C;
            font-size: 0.8em;
        }

        .text-right {
            text-align: right;
        }

        .text-center {
            text-align: center;
        }

        .gray-text {
            color: #DDDDDD;
        }        

        .pale-red-text {
            color: #FFAAAA;
        }        

    </style>
    <title></title>
</head>

<body>
  $HTML_intro
  
  <tr> <!-- Überschriftenzeile -->
    <td colspan=3><table style='width: 100%; text-align: left; background-color: #112233'><tr><td><h1>HM Inventory</h1></td></tr></table></td>
  </tr> <!-- Ende Überschriftenzeile -->
  
  $HTML_ifcs
  $HTML_sep
  $HTML_dvcs
  $HTML_notes
  $HTML_end
</body>
</html>
HEREDOC;
    }

    private function fetchParentId(): int
    {
        return @IPS_GetInstance($this->InstanceID)['ConnectionID'];
    }

    private function isGatewayActive(int $parentId): bool
    {
        if ($parentId === 0) {
            echo 'Gateway is not configured!' . PHP_EOL . PHP_EOL;
            return false;
        }
        $this->SendDebug('Parent', sprintf('%s (#%s)', IPS_GetName($parentId), $parentId), 0);

        if (($this->GetStatus() !== IS_ACTIVE) || !$this->HasActiveParent()) {
            echo 'Instance is not active!' . PHP_EOL . PHP_EOL;
            return false;
        }
        return true;
    }

    private function formatServiceAddresses(array $ParentConfig, string $IP_adr_Homematic): array
    {
        if ($ParentConfig['UseSSL']) {
            $BidCos_RF_Service_adr = sprintf('https://%s:%s', $IP_adr_Homematic, $ParentConfig['RFSSLPort']);
            $BidCos_IP_Service_adr = sprintf('https://%s:%s', $IP_adr_Homematic, $ParentConfig['IPSSLPort']);
        } else {
            $BidCos_RF_Service_adr = sprintf('http://%s:%s', $IP_adr_Homematic, $ParentConfig['RFPort']);
            $BidCos_IP_Service_adr = sprintf('http://%s:%s', $IP_adr_Homematic, $ParentConfig['IPPort']);
        }
        return [$BidCos_RF_Service_adr, $BidCos_IP_Service_adr];
    }

    /**
     * Retrieves the details of a library from a JSON file.
     *
     * @return array The details of the library as an associative array.
     * @throws \JsonException
     */
    private function getLibraryDetails(): array
    {
        $filename = __DIR__ . DIRECTORY_SEPARATOR . '..' . DIRECTORY_SEPARATOR . 'library.json';
        return json_decode(file_get_contents($filename), true, 512, JSON_THROW_ON_ERROR);
    }

    /**
     * Retrieves the version of the module.
     *
     * @return string The version of the module in the format "{major version}.{build number}".
     * @throws \JsonException
     */
    private function getModuleVersion(): string
    {
        $library = $this->getLibraryDetails();
        return sprintf('%s.%s', $library['version'], $library['build']);
    }

    /**
     * Processes HM devices by type and updates the device list, device counter, and error count.
     *
     * @param string  $deviceType   The type of device.
     * @param int     $parentId     The ID of the parent.
     * @param array   $parentConfig The configuration for the parent.
     * @param array  &$hmDevList    The device list.
     * @param array  &$devCounter   The device counter.
     * @param int    &$err          The error count.
     *
     * @return void
     * @throws \JsonException
     */
    private function processHmDevicesByType(
        string $deviceType,
        int $parentId,
        array $parentConfig,
        array &$hmDevList,
        array &$devCounter,
        int &$err
    ): void {
        $deviceHost      = IPS_GetProperty($parentId, 'Host');
        $serviceAddress  = $this->getBidCosServiceAddress($deviceHost, $parentConfig, $deviceType);
        $requestResponse = $this->SendRequestMessage(
            'listDevices',
            [],
            $serviceAddress,
            $parentConfig['UseSSL'],
            $parentConfig['Password'],
            $parentConfig['Username']
        );

        if ($requestResponse->errno === 0) {
            $decodedResponse = php_xmlrpc_decode($requestResponse->value());
            foreach ($decodedResponse as $device) {
                if (($device['PARENT'] === '') && ($device['ADDRESS'] !== 'BidCoS-Wir')) {
                    $devCounter[$deviceType]++;
                }
            }
            $hmDevList = array_merge($hmDevList, $decodedResponse);
        } else {
            $this->SendDebug('Error', sprintf(self::ERROR_MSG, $deviceType), 0);
            $err++;
        }
    }

    /**
     * Retrieves the device lists for a given parent ID and parent configuration.
     *
     * @param int   $parentId     The ID of the parent.
     * @param array $parentConfig The configuration for the parent.
     *
     * @return array The device lists, devCounter, and error value.
     * @throws \JsonException
     */
    private function getDeviceLists(int $parentId, array $parentConfig): array
    {
        $hmDevList  = [];
        $devCounter = [];
        $err        = 0;
        foreach (self::SERVICETYPES as $type) {
            $devCounter[$type] = 0;
            $this->processHmDevicesByType($type, $parentId, $parentConfig, $hmDevList, $devCounter, $err);
        }
        return [$hmDevList, $devCounter, $err];
    }

    private function getBidCosServiceAddress(string $IP_adr_Homematic, array $ParentConfig, string $type): string
    {
        $useSSL = $ParentConfig['UseSSL'] ? 'https://%s:%s' : 'http://%s:%s';
        $port   = $ParentConfig[$type . 'SSLPort'];
        if (!$ParentConfig['UseSSL']) {
            $port = $ParentConfig[$type . 'Port'];
        }
        return sprintf($useSSL, $IP_adr_Homematic, $port);
    }

    private function progressBarInit(): void
    {
        $this->UpdateFormField('ProgressBar', 'maximum', 9);
        $this->UpdateFormField('ProgressBar', 'visible', true);
    }

    private function RegisterProperties(): void
    {
        $this->RegisterPropertyBoolean(self::PROP_ACTIVE, true);
        $this->RegisterPropertyBoolean(self::PROP_SAVEDEVICELISTINVARIABLE, false);
        $this->RegisterPropertyString(self::PROP_OUTPUTFILE, IPS_GetKernelDir() . 'user' . DIRECTORY_SEPARATOR . 'HM_inventory.html');
        $this->RegisterPropertyInteger(self::PROP_SORTORDER, 0);
        $this->RegisterPropertyBoolean(self::PROP_SHOWVIRTUALKEYENTRIES, false);
        $this->RegisterPropertyBoolean(self::PROP_SHOWMAINTENANCEENTRIES, true);
        $this->RegisterPropertyBoolean(self::PROP_SHOWNOTUSEDCHANNELS, true);
        $this->RegisterPropertyBoolean(self::PROP_SHOWLONGIPSDEVICENAMES, false);
        $this->RegisterPropertyBoolean(self::PROP_SHOWHMCONFIGURATORDEVICENAMES, true);
        $this->RegisterPropertyInteger(self::PROP_UPDATEINTERVAL, 0);
    }

    private function RegisterVariables(): void
    {
        if ($this->ReadPropertyBoolean(self::PROP_SAVEDEVICELISTINVARIABLE)) {
            $this->RegisterVariableString('DeviceList', 'Device Liste', '', 1);
        }
    }

    private function SetInstanceStatus(): void
    {
        if ($this->HasActiveParent() && $this->ReadPropertyBoolean(self::PROP_ACTIVE)) {
            $this->SetStatus(IS_ACTIVE);
        } else {
            $this->SetStatus(IS_INACTIVE);
        }
    }

    private function SendRequestMessage(
        string $methodName,
        array $params,
        string $BidCos_Service_adr,
        $UseSSL,
        string $Password,
        string $Username
    ): PhpXmlRpc\Response {
        $xml_BidCos_client = new xmlrpc_client($BidCos_Service_adr);
        if ($UseSSL) {
            $xml_BidCos_client->setSSLVerifyHost(0);
            $xml_BidCos_client->setSSLVerifyPeer(false);
        }
        if ($Password !== '') {
            $xml_BidCos_client->setCredentials($Username, $Password);
        }

        $xml_reqmsg = new xmlrpcmsg($methodName, $params);

        $this->SendDebug(
            'send (xmlrpc)',
            sprintf(
                'send (xmlrpc):%s:%s, params: %s',
                $BidCos_Service_adr,
                $methodName,
                json_encode($params, JSON_THROW_ON_ERROR)
            ),
            0
        );
        return $xml_BidCos_client->send($xml_reqmsg);
    }

    /**
     * Sorts an array of HM addresses in ascending order.
     *
     * The function compares the "HM_address" key of each array element using a case-insensitive comparison.
     * If the "HM_address" is in the format "xxx:yyy" where xxx and yyy are numeric values, the function
     * compares the numeric part of the address as well.
     *
     * @param array $a The first array element to compare.
     * @param array $b The second array element to compare.
     *
     * @return int Returns a negative value if $a is less than $b, 0 if they are equal, and a positive value if $a is greater than $b.
     */
    private static function usort_HM_address(array $a, array $b): int
    {
        $result = strcasecmp($a['HM_address'], $b['HM_address']);

        $a_adr = explode(':', $a['HM_address']);
        $b_adr = explode(':', $b['HM_address']);
        if (count($a_adr) === 2 && count($b_adr) === 2 && strcasecmp($a_adr[0], $b_adr[0]) === 0) {
            $result = (int)($a_adr[1] > $b_adr[1]);
        }

        return $result;
    }

    private static function usort_IPS_dev_name(array $a, array $b): int
    {
        if (($result = strcasecmp($a['IPS_name'], $b['IPS_name'])) === 0) {
            $result = self::usort_HM_address($a, $b);
        }

        return $result;
    }

    private static function usort_HM_device_adr(array $a, array $b): int
    {
        if (($result = strcasecmp($a['HM_device'], $b['HM_device'])) === 0) {
            $result = self::usort_HM_address($a, $b);
        }

        return $result;
    }

    private static function usort_HM_devtype(array $a, array $b): int
    {
        if (($result = strcasecmp($a['HM_devtype'], $b['HM_devtype'])) === 0) {
            $result = self::usort_HM_address($a, $b);
        }

        return $result;
    }

    private static function usort_HM_devname(array $a, array $b): int
    {
        if (($result = strcasecmp($a['HM_devname'], $b['HM_devname'])) === 0) {
            $result = self::usort_HM_address($a, $b);
        }

        return $result;
    }

    private function getHMChannelName($HMAddress, $HMDeviceAddress): string
    {
        $HMScript = 'Name = (xmlrpc.GetObjectByHSSAddress(interfaces.GetAt(0), "' . $HMDeviceAddress . '")).Name();' . PHP_EOL;

        $ScriptReturn  = $this->SendScript($HMAddress, $HMScript);
        $HMChannelName = json_decode($ScriptReturn, true, 512, JSON_THROW_ON_ERROR)['Name'];

        if (!is_string($HMChannelName)) { //Wenn der ChannelName auf HM Seite leer ist, kommt ein leeres Array zurück
            $HMChannelName = '';
        }

        $this->SendDebug(__FUNCTION__, sprintf('HMAddress: %s, HMDeviceAddress: %s -> %s', $HMAddress, $HMDeviceAddress, $HMChannelName), 0);
        return $HMChannelName;
    }

    private function SendScript($HMAddress, $Script): false|string
    {
        $url = 'Script.exe';

        try {
            $HMScriptResult = $this->LoadHMScript($HMAddress, $url, $Script);
            $xml            = @new SimpleXMLElement(mb_convert_encoding($HMScriptResult, 'UTF-8', 'ISO-8859-1'), LIBXML_NOBLANKS + LIBXML_NONET);
        } catch (Exception $exc) {
            trigger_error($exc->getMessage());
        }
        if (isset($xml)) {
            unset($xml->exec, $xml->sessionId, $xml->httpUserAgent);
            return json_encode($xml, JSON_THROW_ON_ERROR);
        }

        return false;
    }


    /**
     * Sets up a cURL handler for making requests to the specified HM address.
     *
     * @param string $HMAddress The CCU address.
     * @param string $url       The URL to send the request to.
     * @param string $HMScript  The payload to send with the request.
     *
     * @return resource|false The cURL handler, or false if the CCU address is not set.
     * @throws \JsonException
     */
    private function setupCurlHandler(string $HMAddress, string $url, string $HMScript): CurlHandle|false
    {
        $ParentId     = @IPS_GetInstance($this->InstanceID)['ConnectionID'];
        $ParentConfig = json_decode(IPS_GetConfiguration($ParentId), true, 512, JSON_THROW_ON_ERROR);
        $scheme       = $ParentConfig['UseSSL'] ? 'https' : 'http';
        $port         = $ParentConfig['UseSSL'] ? $ParentConfig['HSSSLPort'] : $ParentConfig['HSPort'];

        $ch = curl_init(sprintf('%s://%s:%s/%s', $scheme, $HMAddress, $port, $url));

        $header[] = 'Accept: text/plain,text/xml,application/xml,application/xhtml+xml,text/html';
        $header[] = 'Cache-Control: max-age=0';
        $header[] = 'Connection: close';
        $header[] = 'Accept-Charset: UTF-8';
        $header[] = 'Content-type: text/plain;charset="UTF-8"';
        $header[] = 'Expect:';

        if ($ParentConfig['UseSSL']) {
            curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
            curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, false);
        }

        curl_setopt($ch, CURLOPT_HEADER, false);
        curl_setopt($ch, CURLOPT_FAILONERROR, true);
        curl_setopt($ch, CURLOPT_POST, true);
        curl_setopt($ch, CURLOPT_POSTFIELDS, $HMScript);
        curl_setopt($ch, CURLOPT_HTTPHEADER, $header);
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
        curl_setopt($ch, CURLOPT_CONNECTTIMEOUT_MS, 1000);
        curl_setopt($ch, CURLOPT_HTTP_VERSION, CURL_HTTP_VERSION_1_1);
        curl_setopt($ch, CURLOPT_TIMEOUT_MS, 5000);

        if ($ParentConfig['Password'] !== '') {
            curl_setopt($ch, CURLOPT_HTTPAUTH, CURLAUTH_BASIC);
            curl_setopt($ch, CURLOPT_USERPWD, $ParentConfig['Username'] . ':' . $ParentConfig['Password']);
        }

        return $ch;
    }

    /**
     * Loads an HM script from a given HM address using cURL.
     *
     * @param string $HMAddress The CCU address.
     * @param string $url       The URL of the HM script to load.
     * @param string $HMScript  The HM script to load.
     *
     * @return bool|string The result of the cURL execution, or false if the CCU address is not set.
     * @throws \InvalidArgumentException
     * @throws \JsonException
     * @throws \RuntimeException
     */
    private function LoadHMScript(string $HMAddress, string $url, string $HMScript): bool|string
    {
        if ($HMAddress === '') {
            throw new InvalidArgumentException('CCU Address not set.');
        }
        $ch     = $this->setupCurlHandler($HMAddress, $url, $HMScript);
        $result = curl_exec($ch);
        $this->handleCurlError($ch, $result);
        return $result;
    }

    /**
     * Handles CURL errors and throws a RuntimeException if the request fails or if the HTTP response code is greater than or equal to 400.
     *
     * @param \CurlHandle $ch     The CURL handle.
     * @param mixed       $result The result of the CURL request.
     *
     * @return void
     * @throws \RuntimeException If the CURL request fails or if the HTTP response code is greater than or equal to 400.
     */
    private function handleCurlError(CurlHandle $ch, false|string $result): void
    {
        $http_code = curl_getinfo($ch, CURLINFO_HTTP_CODE);
        if (($result === false) || ($http_code >= 400)) {
            throw new RuntimeException('CCU unreachable');
        }
    }
}
