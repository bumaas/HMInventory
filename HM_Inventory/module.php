<?php /** @noinspection CurlSslServerSpoofingInspection */

declare(strict_types=1);

// We need the "xmlrpc" include file
// see https://github.com/gggeek/phpxmlrpc/releases
include_once __DIR__ . '/../libs/phpxmlrpc-4.3.0/lib/xmlrpc.inc';

// Klassendefinition

/** @noinspection AutoloadingIssuesInspection */
class HMInventoryReportCreator extends IPSModule
{
    private const SERVICETYPES = ['RF', 'IP', 'WR'];
    private const ERROR_MSG    = "Can't get any device information from the BidCoS-%s-Service";

    // Some color options for the HTML output
    private const BG_COLOR_GLOBAL         = '#181818';         // Global background color
    private const BG_COLOR_INTERFACE_LIST = '#223344';         // Background color for the interface list
    private const BG_COLOR_HEADLINE       = '#334455';         // Background color for the header line of the device list
    private const BG_COLOR_ODDLINE        = '#181818';         // Background color for the odd lines of the device list
    private const BG_COLOR_EVENLINE       = '#1A2B3C';         // Background color for the even lines of the device list
    private const INVALID_LEVEL           = 65536;

    //property names
    private const PROP_ACTIVE                   = 'active';
    private const PROP_OUTPUTFILE               = 'OutputFile';
    private const PROP_SHOWMAINTENANCEENTRIES   = 'ShowMaintenanceEntries';
    private const PROP_SHOWVIRTUALKEYENTRIES    = 'ShowVirtualKeyEntries';
    private const PROP_SAVEDEVICELISTINVARIABLE = 'SaveDeviceListInVariable';

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

        //Set receive filter to something that will never match
        $this->SetReceiveDataFilter(
            'Dieser Filter sollte niemals greifen, daher ist er etwas länger :-'
        ); //es werden keine Nachrichten vom verbundenen Socket verarbeitet

        if (IPS_GetKernelRunlevel() !== KR_READY) {
            return;
        }

        $this->SetTimerInterval('Update', $this->ReadPropertyInteger('UpdateInterval') * 60 * 1000);

        $this->RegisterVariables();

        $this->SetInstanceStatus();

        $this->SetSummary($this->ReadPropertyString(self::PROP_OUTPUTFILE));
    }

    public function MessageSink($TimeStamp, $SenderID, $Message, $Data): void
    {
        parent::MessageSink($TimeStamp, $SenderID, $Message, $Data);

        if (($Message === IPS_KERNELMESSAGE) && ($Data[0] === KR_READY)) {
            $this->ApplyChanges();
        }
    }

    public function ReceiveData($JSONString): bool
    {
        trigger_error(sprintf('Fatal error: no ReceiveData expected. (%s)', $JSONString));

        return parent::ReceiveData($JSONString);
    }


    /**
     * Die folgenden Funktionen stehen automatisch zur Verfügung, wenn das Modul über die "Module Control" eingefügt wurden.
     * Die Funktionen werden, mit dem selbst eingerichteten Prefix, in PHP und JSON-RPC wie folgt zur Verfügung gestellt:.
     */
    public function CreateReport(): bool
    {
        // originally written by Andreas Bahrdt (HM-Inventory)
        //
        // Public domain
        //
        // Quellen:     http://www.ip-symcon.de/forum/99682-post76.html
        //              https://www.symcon.de/forum/threads/17633-HM_Inventory

        // Anpassung für IPS v 2.5 27.10.2011 by Raketenschnecke
        // Anpassung für IPS v 4.0 16.03.2016 by bumaas
        // Erweiterung für HM-IP und HM-Wired 18.01.2017 by bumaas

        // Get the required data from the BidCos-Services (RF, IP, Wired)

        $ParentId = $this->fetchParentId();
        if (!$this->isGatewayActive($ParentId)) {
            return false;
        }

        $ParentConfig  = json_decode(IPS_GetConfiguration($ParentId), true, 512, JSON_THROW_ON_ERROR);
        $moduleVersion = $this->getModuleVersion();

        $IP_adr_Homematic = IPS_GetProperty($ParentId, 'Host');

        [$BidCos_RF_Service_adr, $BidCos_IP_Service_adr] = $this->formatServiceAddresses($ParentConfig, $IP_adr_Homematic);

        [$hm_dev_list, $dev_counter, $err] = $this->getDeviceLists($ParentId, $ParentConfig);
        if (count($hm_dev_list) === 0) {
            trigger_error("Can't get any device information from the BidCos-Services (Error: $err)", E_USER_ERROR);
        }


        //print_r($hm_dev_list);
        //print_r($dev_counter);

        $this->progressBarInit();
        $progressBarCounter = 0;


        // get all BidCos Interfaces
        $this->UpdateFormField('ProgressBar', 'current', $progressBarCounter++);

        $xml_rtnmsg = $this->SendRequestMessage(
            'listBidcosInterfaces',
            [],
            $BidCos_RF_Service_adr,
            $ParentConfig['UseSSL'],
            $ParentConfig['Password'],
            $ParentConfig['Username']
        );

        if ($xml_rtnmsg->errno === 0) {
            $hm_BidCos_Ifc_list = php_xmlrpc_decode($xml_rtnmsg->value());
            $this->SendDebug('received (BidCos_Ifc_list):', json_encode($hm_BidCos_Ifc_list, JSON_THROW_ON_ERROR), 0);
            //nach 'DEFAULT' sortieren, damit die CCU an erster Stelle steht.
            $default = array_column($hm_BidCos_Ifc_list, 'DEFAULT');
            array_multisort($default, SORT_DESC, $hm_BidCos_Ifc_list);
            //print_r($hm_BidCos_Ifc_list);
        } else {
            $this->SendDebug('Error', "Can't get HM-interface information from the BidCos-RF-Service", 0);
            die("Fatal error: Can't get HM-interface information from the BidCos-Service ($BidCos_RF_Service_adr) - ($xml_rtnmsg->errstr");
        }

        $HM_interface_num           = 0;
        $HM_interface_connected_num = 0;
        $HM_default_interface_no    = 0;

        foreach ($hm_BidCos_Ifc_list as $key => $hm_ifce) {
            $HM_interface_num++;
            if ($hm_ifce['CONNECTED']) {
                $HM_interface_connected_num++;
            }
            if ($hm_ifce['DEFAULT']) {
                $HM_default_interface_no = $key;
            }
        }

        $IPS_device_num     = 0; //Anzahl der IPS Instanzen
        $IPS_HM_channel_num = 0; //Anzahl der verbundenen HM Kanälen (berücksichtigt Mehrfacheinbindungen)
        $HM_module_num      = 0;
        $HM_array           = [];

        // Fill array with all HM-devices found in IP-Symcon
        //
        $this->UpdateFormField('ProgressBar', 'current', $progressBarCounter++);

        foreach (IPS_GetInstanceListByModuleID('{EE4A81C6-5C90-4DB7-AD2F-F6BBD521412E}') as $id) {
            //first check if the device is assigned to the right gateway
            if ($ParentId !== IPS_GetInstance($id)['ConnectionID']) {
                continue;
            }
            $HM_module_num++;
            $IPS_device_num++;
            $IPS_HM_channel_already_assigned = false;
            $HM_address                      = IPS_GetProperty($id, 'Address');
            $this->SendDebug('hm_device', 'ID: ' . $id . ', Address: ' . $HM_address, 0);

            $NeedlePos = strpos($HM_address, ':');
            if ($NeedlePos) {
                $HM_Par_address = substr($HM_address, 0, $NeedlePos);
            } else {
                echo 'HM address (' . $HM_address . ') of id ' . $id . ' is invalid.' . PHP_EOL;
                continue;
            }
            $HM_device     = '-';
            $HM_devname    = '-';
            $HM_FWversion  = ' ';
            $HM_Interface  = '';
            $HM_Roaming    = ' ';
            $HM_devtype    = '-';
            $HM_direction  = '-';
            $HM_AES_active = '-';
            $hm_chld_dev   = null;
            $hm_par_dev    = null;

            foreach ($hm_dev_list as $hm_dev) {
                if ($hm_dev['ADDRESS'] === $HM_address) {
                    $hm_chld_dev = $hm_dev;
                }
                if ($hm_dev['ADDRESS'] === $HM_Par_address) {
                    $hm_par_dev = $hm_dev;
                }
                if ($hm_chld_dev !== null) {
                    if (isset($hm_dev['PARENT_TYPE'])) {
                        $HM_device = $hm_dev['PARENT_TYPE'];
                    }
                    if ($this->ReadPropertyBoolean('ShowHMConfiguratorDeviceNames')) {
                        $HM_devname = $this->getHMChannelName($IP_adr_Homematic, $hm_dev['ADDRESS']);
                    }

                    $device_info = $this->extractDeviceInfo($hm_par_dev);

                    $HM_FWversion = $device_info['FIRMWARE'];
                    $HM_Interface = $device_info['INTERFACE'];
                    $HM_Roaming   = $device_info['ROAMING'];

                    $HM_devtype = $hm_dev['TYPE'];
                    if (isset($hm_dev['DIRECTION'])) {
                        if ($hm_dev['DIRECTION'] === 1) {
                            $HM_direction = 'TX';
                        } elseif ($hm_dev['DIRECTION'] === 2) {
                            $HM_direction = 'RX';
                        }
                    }
                    if (isset($hm_dev['AES_ACTIVE']) && ($hm_dev['AES_ACTIVE'] !== 0)) {
                        $HM_AES_active = '+';
                    }
                    break;
                }
            }

            if ($HM_address !== '') {
                foreach ($HM_array as &$HM_dev) {
                    if ($HM_dev['HM_address'] === $HM_address) {
                        $HM_dev['IPS_HM_d_assgnd']       = true;
                        $IPS_HM_channel_already_assigned = true;
                        break;
                    }
                }
                unset($HM_dev);
                if (!$IPS_HM_channel_already_assigned) {
                    $IPS_HM_channel_num++;
                }
            }

            if ($this->ReadPropertyBoolean('ShowLongIPSDeviceNames')) {
                $IPS_name = IPS_GetLocation($id);
            } else {
                $IPS_name = IPS_GetName($id);
            }

            $HM_array[] = [
                'IPS_occ'         => $HM_module_num,
                'IPS_id'          => $id,
                'IPS_name'        => $IPS_name,
                'IPS_HM_d_assgnd' => $IPS_HM_channel_already_assigned,
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

        // Add HM_devices known by BidCos but not present in IP-Symcon
        //
        $this->UpdateFormField('ProgressBar', 'current', $progressBarCounter++);

        if ($this->ReadPropertyBoolean('ShowNotUsedChannels')) {
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

        // Request tx/rx RF-levels from BidCos-RF-Service
        //
        $this->UpdateFormField('ProgressBar', 'current', $progressBarCounter++);

        $xml_rtnmsg = $this->SendRequestMessage(
            'rssiInfo',
            [],
            $BidCos_RF_Service_adr,
            $ParentConfig['UseSSL'],
            $ParentConfig['Password'],
            $ParentConfig['Username']
        );

        $hm_lvl_list = [];
        if ($xml_rtnmsg->errno === 0) {
            $hm_lvl_list = php_xmlrpc_decode($xml_rtnmsg->value());
            //print_r($hm_lvl_list);
        } else {
            echo "Warning: Can't get RF-level information from the BidCos-Service ($BidCos_RF_Service_adr) &nbsp&nbsp&nbsp - &nbsp&nbsp&nbsp ($xml_rtnmsg->errstr)<br>\n";
        }

        // Add tx/rx RF-levels for each device/interface
        //
        $this->UpdateFormField('ProgressBar', 'current', $progressBarCounter++);

        if (is_array($hm_lvl_list)) {
            foreach ($HM_array as &$HM_dev) {
                $hm_adr = explode(':', $HM_dev['HM_address']);
                if (isset($hm_adr[0]) && array_key_exists($hm_adr[0], $hm_lvl_list)) {
                    $HM_lvl_array  = [];
                    $hm_levels     = $hm_lvl_list[$hm_adr[0]];
                    $best_lvl_ifce = -1;
                    $ifce_no       = 0;
                    foreach ($hm_BidCos_Ifc_list as $hm_ifce) {
                        if ($hm_ifce['CONNECTED']) {
                            if (array_key_exists($hm_ifce['ADDRESS'], $hm_levels)) {
                                $HM_lvl_array[] = [
                                    $hm_levels[$hm_ifce['ADDRESS']][0],
                                    $hm_levels[$hm_ifce['ADDRESS']][1],
                                    $HM_dev['HM_Interface'] === $hm_ifce['ADDRESS'],
                                    false
                                ];
                                if ($hm_levels[$hm_ifce['ADDRESS']][1] !== self::INVALID_LEVEL) {
                                    if ($best_lvl_ifce === -1) {
                                        $best_lvl_ifce = $ifce_no;
                                    } elseif ($HM_lvl_array[$best_lvl_ifce][1] < $hm_levels[$hm_ifce['ADDRESS']][1]) {
                                        $best_lvl_ifce = $ifce_no;
                                    }
                                }
                            } else {
                                $HM_lvl_array[] = [self::INVALID_LEVEL, self::INVALID_LEVEL, false, false];
                            }
                            $ifce_no++;
                        }
                    }
                    if ($best_lvl_ifce !== -1) {
                        $best_lvl = $HM_lvl_array[$best_lvl_ifce][1];
                        foreach ($HM_lvl_array as &$hm_lvl) {
                            $hm_lvl[3] = ($hm_lvl[1] === $best_lvl);
                        }
                        unset($hm_lvl);
                    }
                    $HM_dev['HM_levels'] = $HM_lvl_array;
                }
            }
            unset($HM_dev);
        }

        // Request tx/rx RF-levels from BidCos-IP-Service
        //
        $this->UpdateFormField('ProgressBar', 'current', $progressBarCounter++);

        usort($HM_array, self::usort_HM_address(...));

        $previous_hm_adr    = '';
        $previous_hm_levels = [];
        foreach ($HM_array as &$HM_dev) {
            $hm_adr = explode(':', $HM_dev['HM_address']);
            if (isset($hm_adr[0]) && strlen($hm_adr[0]) !== 14) {
                $previous_hm_adr = '';
                continue;
            }
            if ($hm_adr[0] !== $previous_hm_adr) {
                $params     = [
                    new xmlrpcval($hm_adr[0] . ':0', 'string'),
                    new xmlrpcval('VALUES', 'string')
                ];
                $xml_rtnmsg = $this->SendRequestMessage(
                    'getParamset',
                    $params,
                    $BidCos_IP_Service_adr,
                    $ParentConfig['UseSSL'],
                    $ParentConfig['Password'],
                    $ParentConfig['Username']
                );

                if ($xml_rtnmsg->errno === 0) {
                    $HM_ParamSet = php_xmlrpc_decode($xml_rtnmsg->value());
                    //print_r($HM_ParamSet);
                    $HM_dev['HM_levels'][$HM_default_interface_no][0] = $HM_ParamSet['RSSI_PEER'] ?? self::INVALID_LEVEL;
                    $HM_dev['HM_levels'][$HM_default_interface_no][1] = $HM_ParamSet['RSSI_DEVICE'] ?? self::INVALID_LEVEL;
                    $HM_dev['HM_levels'][$HM_default_interface_no][2] = false; //??
                    $HM_dev['HM_levels'][$HM_default_interface_no][3] = false; //best level

                    $previous_hm_levels = $HM_dev['HM_levels'];
                }
            } else {
                $HM_dev['HM_levels'] = $previous_hm_levels;
            }
        }
        unset($HM_dev);


        //delete the Maintenance Channels and Virtual Keys if required
        foreach ($HM_array as $key => $HM_dev) {
            if ($HM_dev['HM_devtype'] === 'MAINTENANCE' && !$this->ReadPropertyBoolean(self::PROP_SHOWMAINTENANCEENTRIES)) {
                unset ($HM_array[$key]);
            }

            if ($HM_dev['HM_devtype'] === 'VIRTUAL_KEY' && !$this->ReadPropertyBoolean(self::PROP_SHOWVIRTUALKEYENTRIES)) {
                unset ($HM_array[$key]);
            }
        }


        // Sort device list
        //
        $SortOrder = $this->ReadPropertyInteger('SortOrder');
        switch ($SortOrder) {
            case 3: // by IPS-dev_name
                usort($HM_array, self::usort_IPS_dev_name(...));
                break;
            case 0: // by HM-address
                usort($HM_array, self::usort_HM_address(...));
                break;
            case 1: // by HM-device
                usort($HM_array, self::usort_HM_device_adr(...));
                break;
            case 2: //by HM-type
                usort($HM_array, self::usort_HM_devtype(...));
                break;
            case 4: //by HM-device-name
                usort($HM_array, self::usort_HM_devname(...));
                break;
            default:
                trigger_error('Unknown SortOrder: ' . $SortOrder);
        }

        //print_r($HM_array);

        if ($this->ReadPropertyBoolean(self::PROP_SAVEDEVICELISTINVARIABLE)) {
            //SetValueString($this->GetIDForIdent('DeviceList'), json_encode($HM_array));
            $this->SetValue('DeviceList', json_encode($HM_array, JSON_THROW_ON_ERROR)); //array in String variable speichern
        }

        // Generate HTML output code

        $this->UpdateFormField('ProgressBar', 'current', $progressBarCounter);

        $HTML_intro = "<table width='100%' border='0' align='center' bgcolor=" . self::BG_COLOR_GLOBAL . '>';

        $HM_inventory_str = sprintf('<b>HM Inventory (%s) </b><b>&nbsp found at %s</b>', $moduleVersion, date('d.m.Y H:i:s'));
        $HM_interface_str = sprintf(
            '%s HomeMatic interfaces (%s connected) with %s HM-RF devices, %s HM-wired devices and %s HmIP devices',
            $HM_interface_num,
            $HM_interface_connected_num,
            $dev_counter['RF'],
            $dev_counter['WR'],
            $dev_counter['IP']
        );
        $IPS_instance_str = sprintf('%s IPS instances (connected to %s HM channels)', $IPS_device_num, $IPS_HM_channel_num);

        $HTML_ifcs = "<tr style='vertical-align: top'>";
        $HTML_ifcs .= "<td><table style='text-align: left;font-size: large; color: #99AABB'>";
        $HTML_ifcs .= $this->generateTableRow('large', '#99AABB', $HM_inventory_str);
        $HTML_ifcs .= $this->generateTableRow('small', '#CCCCCC', $HM_interface_str);
        $HTML_ifcs .= $this->generateTableRow('small', '#CCCCCC', $IPS_instance_str);
        $HTML_ifcs .= '</table></td>';
        $HTML_ifcs .= "<td style='vertical-align: top'>&nbsp;</td>";
        $HTML_ifcs .= "<td style='width: 40%; vertical-align: bottom;'><table style='width: 100%; text-align: right; background-color: "
                      . self::BG_COLOR_INTERFACE_LIST . '\'>';

        //print_r($hm_BidCos_Ifc_list);
        foreach ($hm_BidCos_Ifc_list as $hm_ifce) {
            $HTML_ifcs .= $this->formatInterfaceRow($hm_ifce);
        }
        $HTML_ifcs .= '</table></td></tr>';

        $HTML_sep = '<tr><td colspan=3><table style="width: 100%; text-align: left"> <hr><tr><td> </td></tr></table></td></tr>';

        $dthdr_td_b   = '<td style="font-size: small; color: #EEEEEE"><b>';
        $dthdr_td_b_r = '<td style="text-align: right; font-size: small; color: #EEEEEE"><b>';
        $dthdr_td_e   = '</b></td>';
        $dthdr_td_eb  = $dthdr_td_e . $dthdr_td_b;
        $HTML_dvcs    = '<tr><td colspan=3><table style="width: 100%; text-align: left">';
        $HTML_dvcs    .= '<tr bgcolor=' . self::BG_COLOR_HEADLINE . '>';
        $HTML_dvcs    .= $dthdr_td_b_r . '&nbsp##&nbsp' . $dthdr_td_eb . 'IPS ID' . $dthdr_td_eb . 'IPS device name&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp'
                         . $dthdr_td_eb . 'HM address' . $dthdr_td_e;
        if ($this->ReadPropertyBoolean('ShowHMConfiguratorDeviceNames')) {
            $HTML_dvcs .= $dthdr_td_b . 'HM device name' . $dthdr_td_e;
        }
        $HTML_dvcs .= $dthdr_td_b . 'HM device type' . $dthdr_td_eb . 'Fw.' . $dthdr_td_eb . 'HM channel type' . $dthdr_td_eb . 'Dir.' . $dthdr_td_eb
                      . 'AES' . $dthdr_td_e;
        $HTML_dvcs .= '<td style="width: 2%; text-align: center; color: #EEEEEE; font-size: medium">Roa- ming</td>';
        foreach ($hm_BidCos_Ifc_list as $hm_ifce) {
            if ($hm_ifce['CONNECTED']) {
                $HTML_dvcs .= '<td style="width: 6%; text-align: center; color: #EEEEEE; font-size: small">' . $hm_ifce['ADDRESS']
                              . ' tx/rx&nbsp(db&micro;V)' . '</td>';
            }
        }
        $HTML_dvcs .= '</tr>';

        $entry_no        = 0;
        $previous_hm_adr = '';
        foreach ($HM_array as $HM_dev) {
            $hm_adr = explode(':', $HM_dev['HM_address']);
            if ($hm_adr[0] === $previous_hm_adr) {
                $same_device = true;
            } else {
                $same_device     = false;
                $previous_hm_adr = $hm_adr[0];
            }
            $font_tag      = '<font size="2" color=' . (($HM_dev['IPS_HM_d_assgnd'] === false) ? '#DDDDDD' : '#FFAAAA') . '>';
            $dtdvc_td_b    = '<td>' . $font_tag;
            $dtdvc_td_ar_b = '<td style="text-align: right">' . $font_tag;
            $dtdvc_td_ac_b = '<td style="text-align: center">' . $font_tag;
            $dtdvc_td_e    = '</font></td>';
            $dtdvc_td_eb   = $dtdvc_td_e . $dtdvc_td_b;
            if (($entry_no++ % 2) === 0) {
                $r_bgcolor = self::BG_COLOR_ODDLINE;
            } else {
                $r_bgcolor = self::BG_COLOR_EVENLINE;
            }
            $HTML_dvcs .= '<tr bgcolor=' . $r_bgcolor . '>' . $dtdvc_td_ar_b . $entry_no . '&nbsp&nbsp' . $dtdvc_td_eb;
            $HTML_dvcs .= $HM_dev['IPS_id'] . $dtdvc_td_eb . mb_convert_encoding($HM_dev['IPS_name'], 'ISO-8859-1', 'UTF-8') . $dtdvc_td_eb
                          . $HM_dev['HM_address'] . $dtdvc_td_eb;
            if ($this->ReadPropertyBoolean('ShowHMConfiguratorDeviceNames')) {
                $HTML_dvcs .= mb_convert_encoding($HM_dev['HM_devname'], 'ISO-8859-1', 'UTF-8') . $dtdvc_td_eb;
            }
            if (!$same_device) {
                $HTML_dvcs .= $HM_dev['HM_device'] . $dtdvc_td_eb . $HM_dev['HM_FWversion'] . $dtdvc_td_eb;
            } else {
                $HTML_dvcs .= $dtdvc_td_eb . $dtdvc_td_eb;
            }
            $HTML_dvcs .= $HM_dev['HM_devtype'] . $dtdvc_td_eb . $HM_dev['HM_direction'] . $dtdvc_td_e . $dtdvc_td_ac_b . $HM_dev['HM_AES_active']
                          . $dtdvc_td_e . $dtdvc_td_ac_b;

            if (!$same_device) {
                $HTML_dvcs .= $HM_dev['HM_Roaming'] . $dtdvc_td_e;
                //print_r($HM_dev);

                if (isset($HM_dev['HM_levels'])) {
                    foreach ($HM_dev['HM_levels'] as $lci => $lciValue) {
                        if (isset($HM_dev['HM_levels'][$lci])) {
                            // Interface with best levels gets different color
                            if (!isset($lciValue[3])) {
                                echo $HM_dev['HM_device'] . PHP_EOL;
                            }
                            if ($lciValue[3]) {
                                if (($HM_dev['HM_Roaming'] === '+') || $lciValue[2]) {
                                    $lvl_strg_color = '<p style="color: #DDDD66">';
                                } else {
                                    $lvl_strg_color = '<p style="color: #FFFF88">';
                                }
                            } else {
                                $lvl_strg_color = '<p style="color: =#DDDDDD">';
                            }

                            [$rx_strg, $tx_strg] = $this->getRxTxLevelString($lciValue[0], $lciValue[1]);

                            if (($HM_dev['HM_Roaming'] === '+') || $lciValue[2]) {
                                $lvl_strg = sprintf(
                                    '%s<ins>%s &#047 %s</ins></p>',
                                    $lvl_strg_color,
                                    $rx_strg,
                                    $tx_strg
                                );
                            } else {
                                $lvl_strg = sprintf(
                                    '%s%s &#047 %s</p>',
                                    $lvl_strg_color,
                                    $rx_strg,
                                    $tx_strg
                                );
                            }

                            $HTML_dvcs .= $dtdvc_td_ac_b . $lvl_strg . $dtdvc_td_e;
                        } else {
                            $HTML_dvcs .= '<td> </td>';
                        }
                    }
                }
            } else {
                $HTML_dvcs .= $dtdvc_td_e . '<td> </td>';
                $HTML_dvcs .= str_repeat('<td> </td>', $HM_interface_connected_num);
            }
        }

        if ($HM_module_num === 0) {
            $HTML_dvcs .= '<tr><td colspan=20 style="text-align: center; color: #DDDDDD; font-size: large"><br/>No HomeMatic devices found!</td></tr>';
        }

        $HTML_dvcs .= '</td></tr>';

        // Some comments
        //
        $HTML_notes = '<tr><td colspan=20><table style="width: 100%; text-align: left; color: #666666"><hr><tr><td> </td></tr></table></td></tr>';
        $HTML_notes .= '<tr><td colspan=20><table style="width: 100%; text-align: left; font-size:medium; color: #DDDDDD"><tr><td>Notes:</td></tr>';
        $HTML_notes .= '<tr><td style="font-size: smaller; color: #DDDDDD"><ol>';

        $notes = [
            'Interfaces: bold letters indicate the default BidCos-Interface.',
            'Level-pairs: the left value is showing the last signal level received by the device from the interface, while the right value is showing the last signal level received by the interface from the device.',
            'Level-pairs: underlined letters of the level-pair indicate the BidCos-Interface associated with the device (or all interfaces when Roaming is enabled for the device).',
            'Level-pairs: the yellow level-pair indicates the BidCos-Interface with best signal quality.',
            'Devices without level-pairs haven\'t sent/received anything since last start of the BidCos-service or are wired.',
            'BidCos channels assigned to more than one IPS-device are shown in red.'
        ];
        foreach ($notes as $note) {
            $HTML_notes .= $this->createHtmlElement($note, 'li');
        }
        $HTML_notes .= '</ol></td></tr>';
        $HTML_notes .= '</table></td></tr>';

        $HTML_end = '</table>';

        $this->UpdateFormField('ProgressBar', 'visible', false);

        // Output the results
        $OutputFileName = $this->ReadPropertyString(self::PROP_OUTPUTFILE);
        if ($OutputFileName) {
            $HTML_file = @fopen($OutputFileName, 'wb');
            if ($HTML_file === false) {
                echo sprintf('File "%s" not writable!' . PHP_EOL . PHP_EOL, $OutputFileName);
                return false;
            }

            // Use single fwrite call
            $htmlContent = $this->getHtmlContent($HTML_intro, $HTML_ifcs, $HTML_sep, $HTML_dvcs, $HTML_notes, $HTML_end);
            fwrite($HTML_file, $htmlContent);

            return fclose($HTML_file);
        }

        return false;
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

    private function createHtmlElement($content, $containerTag, $props = ''): string
    {
        return "<$containerTag $props>$content</$containerTag>";
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
            </tr>";
    }

    private function getRxTxLevelString($rx_lvl, $tx_lvl): array
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
    html,body {font-family:Arial,Helvetica,sans-serif;font-size:12px;background-color:#000000;color:#dddddd;}
  </style><title></title>
</head>
<body>
  $HTML_intro
  <tr>
    <td colspan=3><table style='width: 100%; text-align: left; background-color: #112233'><tr><td><h1>HM inventory</h1></td></tr></table></td>
  </tr>
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
        $this->RegisterPropertyBoolean('SaveHMArrayInVariable', false);
        $path = 'user' . DIRECTORY_SEPARATOR . 'HM_inventory.html';
        if (file_exists(IPS_GetKernelDir() . 'webfront')) {
            $path = 'webfront\\' . $path;
        }
        $this->RegisterPropertyString(self::PROP_OUTPUTFILE, IPS_GetKernelDir() . $path);
        $this->RegisterPropertyInteger('SortOrder', 0);
        $this->RegisterPropertyBoolean(self::PROP_SHOWVIRTUALKEYENTRIES, false);
        $this->RegisterPropertyBoolean(self::PROP_SHOWMAINTENANCEENTRIES, true);
        $this->RegisterPropertyBoolean('ShowNotUsedChannels', true);
        $this->RegisterPropertyBoolean('ShowLongIPSDeviceNames', false);
        $this->RegisterPropertyBoolean('ShowHMConfiguratorDeviceNames', true);
        $this->RegisterPropertyInteger('UpdateInterval', 0);
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

        if (!is_string($HMChannelName)) { //Wenn der ChannelName auf HM Seite leer ist, dann kommt ein leeres Array zurück
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
        curl_close($ch);
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
