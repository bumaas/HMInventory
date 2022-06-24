<?php /** @noinspection CurlSslServerSpoofingInspection */

declare(strict_types=1);

// We need the "xmlrpc" include file
// see https://github.com/gggeek/phpxmlrpc/releases
include_once __DIR__ . '/../libs/phpxmlrpc-4.3.0/lib/xmlrpc.inc';

// Klassendefinition

/** @noinspection AutoloadingIssuesInspection */
class HMInventoryReportCreator extends IPSModule
{
    // Some color options for the HTML output
    private const BG_COLOR_GLOBAL         = '#181818';         // Global background color
    private const BG_COLOR_INTERFACE_LIST = '#223344';         // Background color for the interface list
    private const BG_COLOR_HEADLINE       = '#334455';         // Background color for the header line of the device list
    private const BG_COLOR_ODDLINE        = '#181818';         // Background color for the odd lines of the device list
    private const BG_COLOR_EVENLINE       = '#1A2B3C';         // Background color for the even lines of the device list

    //property names
    private const PROP_OUTPUTFILE = 'OutputFile';
    private const PROP_SHOWMAINTENANCEENTRIES = 'ShowMaintenanceEntries';
    private const PROP_SHOWVIRTUALKEYENTRIES = 'ShowVirtualKeyEntries';
    private const PROP_SAVEDEVICELISTINVARIABLE = 'SaveDeviceListInVariable';

    // Überschreibt die interne IPS_Create($id) Funktion
    public function Create()
    {
        // Diese Zeile nicht löschen.
        parent::Create();

        $this->RegisterProperties();

        $this->RegisterTimer('Update', 0, 'HMI_CreateReport(' . $this->InstanceID . ');');

        //we will wait until the kernel is ready
        $this->RegisterMessage(0, IPS_KERNELMESSAGE);

    }

    public function ApplyChanges()
    {
        //Never delete this line!
        parent::ApplyChanges();

        //Set receive filter to something that will never match
        $this->SetReceiveDataFilter('Dieser Filter sollte niemals greifen, daher ist er etwas länger :-'); //es werden keine Nachrichten vom verbundenen Socket verarbeitet

        if (IPS_GetKernelRunlevel() !== KR_READY) {
            return;
        }

        $this->SetTimerInterval('Update', $this->ReadPropertyInteger('UpdateInterval') * 60 * 1000);

        $this->RegisterVariables();

        $this->SetInstanceStatus();

        $this->SetSummary($this->ReadPropertyString(self::PROP_OUTPUTFILE));
    }

    public function MessageSink($TimeStamp, $SenderID, $Message, $Data)
    {
        parent::MessageSink($TimeStamp, $SenderID, $Message, $Data);

        if (($Message === IPS_KERNELMESSAGE) && ($Data[0] === KR_READY)) {
            $this->ApplyChanges();
        }
    }

    public function ReceiveData($JSONString)
    {
        trigger_error(sprintf ('Fatal error: no ReceiveData expected. (%s)', $JSONString));

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

        $ParentId = @IPS_GetInstance($this->InstanceID)['ConnectionID'];
        $this->SendDebug('Parent', sprintf('%s (#%s)', IPS_GetName($ParentId), $ParentId), 0);
        if ($ParentId === 0) {
            echo 'Gateway is not configured!' . PHP_EOL . PHP_EOL;
            return false;
        }

        if (($this->GetStatus() !== IS_ACTIVE) || !$this->HasActiveParent()) {
            echo 'Instance is not active!' . PHP_EOL . PHP_EOL;
            return false;
        }

        $ParentConfig = json_decode(IPS_GetConfiguration($ParentId), true);

        $IP_adr_Homematic = IPS_GetProperty($ParentId, 'Host');

        if ($ParentConfig['UseSSL']) {
            $BidCos_Wired_Service_adr = sprintf('https://%s:%s', $IP_adr_Homematic, $ParentConfig['WRSSLPort']);
            $BidCos_RF_Service_adr    = sprintf('https://%s:%s', $IP_adr_Homematic, $ParentConfig['RFSSLPort']);
            $BidCos_IP_Service_adr    = sprintf('https://%s:%s', $IP_adr_Homematic, $ParentConfig['IPSSLPort']);
        } else {
            $BidCos_Wired_Service_adr = sprintf('http://%s:%s', $IP_adr_Homematic, $ParentConfig['WRPort']);
            $BidCos_RF_Service_adr    = sprintf('http://%s:%s', $IP_adr_Homematic, $ParentConfig['RFPort']);
            $BidCos_IP_Service_adr    = sprintf('http://%s:%s', $IP_adr_Homematic, $ParentConfig['IPPort']);
        }

        $filename      = __DIR__ . DIRECTORY_SEPARATOR . '..' . DIRECTORY_SEPARATOR . 'library.json';
        $library       = json_decode(file_get_contents($filename), true);
        $moduleVersion = sprintf('%s.%s', $library['version'], $library['build']);


        $hm_RF_dev_list                = [];
        $hm_RF_parent_devices_count    = 0;
        $hm_IP_dev_list                = [];
        $hm_IP_parent_devices_count    = 0;
        $hm_Wired_dev_list             = [];
        $hm_Wired_parent_devices_count = 0;
        $err                           = 0;

        //print_r($this->LoadHMScript($IP_adr_Homematic, $BidCos_RF_Service_adr, 'listDevices'));

        // get the RF devices
        $xml_rtnmsg = $this->SendRequestMessage('listDevices', [], $BidCos_RF_Service_adr, $ParentConfig['UseSSL'], $ParentConfig['Password'], $ParentConfig['Username']);
        if ($xml_rtnmsg->errno === 0) {
            $hm_RF_dev_list = php_xmlrpc_decode($xml_rtnmsg->value());
            $this->SendDebug('received (RF_dev_list):', json_encode($hm_RF_dev_list), 0);
            foreach ($hm_RF_dev_list as $device) {
                if (($device['PARENT'] === '') && ($device['ADDRESS'] !== 'BidCoS-RF')) {
                    //echo sprintf('%s'.PHP_EOL, $device['ADDRESS']);
                    $hm_RF_parent_devices_count++;
                }
                //echo $device
            }
            //print_r($hm_RF_dev_list);
        } else {
            $this->SendDebug('Error', "Can't get any device information from the BidCoS-RF-Service", 0);
            $err++;
        }

        // get the IP devices
        $xml_rtnmsg = $this->SendRequestMessage('listDevices', [], $BidCos_IP_Service_adr, $ParentConfig['UseSSL'], $ParentConfig['Password'], $ParentConfig['Username']);
        if ($xml_rtnmsg->errno === 0) {
            $hm_IP_dev_list = php_xmlrpc_decode($xml_rtnmsg->value());
            $this->SendDebug('received (IP_dev_list):', json_encode($hm_IP_dev_list), 0);
            foreach ($hm_IP_dev_list as $device) {
                if ($device['PARENT'] === '') {
                    //echo sprintf('%s' . PHP_EOL, $device['ADDRESS']);
                    $hm_IP_parent_devices_count++;
                }
            }
            //print_r($hm_IP_dev_list);
        } else {
            $this->SendDebug('Error', "Can't get any device information from the BidCos-IP-Service", 0);
            $err += 2;
        }

        // get the Wired devices
        $xml_rtnmsg = $this->SendRequestMessage('listDevices', [], $BidCos_Wired_Service_adr, $ParentConfig['UseSSL'], $ParentConfig['Password'], $ParentConfig['Username']);
        if ($xml_rtnmsg->errno === 0) {
            $hm_Wired_dev_list = php_xmlrpc_decode($xml_rtnmsg->value());
            $this->SendDebug('received (Wired_dev_list):', json_encode($hm_Wired_dev_list), 0);
            foreach ($hm_Wired_dev_list as $device) {
                if (($device['PARENT'] === '') && ($device['ADDRESS'] !== 'BidCoS-Wir')) {
                    //echo sprintf('%s' . PHP_EOL, $device['ADDRESS']);
                    $hm_Wired_parent_devices_count++;
                }
            }
            //print_r($hm_Wired_dev_list);
        } else {
            $this->SendDebug('Error', "Can't get any device information from the BidCos-Wired-Service", 0);
            $err += 4;
        }

        // merge all devices
        $hm_dev_list = array_merge($hm_RF_dev_list, $hm_IP_dev_list, $hm_Wired_dev_list);
        if (count($hm_dev_list) === 0) {
            trigger_error("Can't get any device information from the BidCos-Services (Error: $err)", E_USER_ERROR);
        }
        //print_r($hm_dev_list);

        $progressBarCounter = 0;
        $this->UpdateFormField('ProgressBar', 'maximum', 9);
        $this->UpdateFormField('ProgressBar', 'visible', true);


        // get all BidCos Interfaces
        $this->UpdateFormField('ProgressBar', 'current', $progressBarCounter++);

        $xml_rtnmsg = $this->SendRequestMessage('listBidcosInterfaces', [], $BidCos_RF_Service_adr, $ParentConfig['UseSSL'], $ParentConfig['Password'], $ParentConfig['Username']);

        if ($xml_rtnmsg->errno === 0) {
            $hm_BidCos_Ifc_list = php_xmlrpc_decode($xml_rtnmsg->value());
            $this->SendDebug('received (BidCos_Ifc_list):', json_encode($hm_BidCos_Ifc_list), 0);
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

            if ((float) IPS_GetKernelVersion() < 5.6) {
                set_time_limit(60); //Abfragen dauern manchmal länger als 30 Sekunden
            }

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
                        $HM_devname = $this->GetHMChannelName($IP_adr_Homematic, $hm_dev['ADDRESS']);
                    }
                    if ($hm_par_dev !== null) {
                        $HM_FWversion = $hm_par_dev['FIRMWARE'];
                        $HM_Interface = $hm_par_dev['INTERFACE'] ?? '';
                        if (isset($hm_par_dev['ROAMING']) && $hm_par_dev['ROAMING']) {
                            $HM_Roaming = '+';
                        } else {
                            $HM_Roaming = '-';
                        }
                    }

                    $HM_devtype = $hm_dev['TYPE'];
                    if (isset($hm_dev['DIRECTION'])) {
                        if ($hm_dev['DIRECTION'] === 1) {
                            $HM_direction = 'TX';
                        } elseif ($hm_dev['DIRECTION'] === 2) {
                            $HM_direction = 'RX';
                        }
                    } else {
                        $HM_direction = '-';
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
                        $HM_FWversion = '';
                        $HM_Interface = '';
                        $HM_Roaming   = '';
                        $HM_module_num++;
                        $HM_device  = $hm_chld_dev['PARENT_TYPE'];
                        $HM_devname = '-';
                        if ($this->ReadPropertyBoolean('ShowHMConfiguratorDeviceNames')) {
                            $HM_devname = $this->GetHMChannelName($IP_adr_Homematic, $HM_address);
                        }

                        if ($hm_par_dev !== null) {
                            $HM_FWversion = $hm_par_dev['FIRMWARE'];
                            $HM_Interface = $hm_par_dev['INTERFACE'] ?? '';
                            if (isset($hm_par_dev['ROAMING']) && $hm_par_dev['ROAMING']) {
                                $HM_Roaming = '+';
                            } else {
                                $HM_Roaming = '-';
                            }
                        }

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

        $xml_rtnmsg = $this->SendRequestMessage('rssiInfo', [], $BidCos_RF_Service_adr, $ParentConfig['UseSSL'], $ParentConfig['Password'], $ParentConfig['Username']);

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
                                if ($hm_levels[$hm_ifce['ADDRESS']][1] !== 65536) {
                                    if ($best_lvl_ifce === -1) {
                                        $best_lvl_ifce = $ifce_no;
                                    } elseif ($HM_lvl_array[$best_lvl_ifce][1] < $hm_levels[$hm_ifce['ADDRESS']][1]) {
                                        $best_lvl_ifce = $ifce_no;
                                    }
                                }
                            } else {
                                $HM_lvl_array[] = [65536, 65536, false, false];
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

        usort($HM_array, 'self::usort_HM_address');

        $previous_hm_adr    = '';
        $previous_hm_levels = [];
        foreach ($HM_array as &$HM_dev) {
            $hm_adr = explode(':', $HM_dev['HM_address']);
            if (isset($hm_adr[0]) && strlen($hm_adr[0]) !== 14) {
                $previous_hm_adr = '';
                continue;
            }
            if ($hm_adr[0] !== $previous_hm_adr) {
                $params = [
                    new xmlrpcval($hm_adr[0] . ':0', 'string'),
                    new xmlrpcval('VALUES', 'string')
                ];
                $xml_rtnmsg = $this->SendRequestMessage('getParamset', $params, $BidCos_IP_Service_adr, $ParentConfig['UseSSL'], $ParentConfig['Password'], $ParentConfig['Username']);

                if ($xml_rtnmsg->errno === 0) {
                    $HM_ParamSet = php_xmlrpc_decode($xml_rtnmsg->value());
                    //print_r($HM_ParamSet);
                    $HM_dev['HM_levels'][$HM_default_interface_no][0] = $HM_ParamSet['RSSI_PEER'] ?? 65536;
                    $HM_dev['HM_levels'][$HM_default_interface_no][1] = $HM_ParamSet['RSSI_DEVICE'] ?? 65536;
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
        foreach ($HM_array as $key=> $HM_dev){
            if ($HM_dev['HM_devtype'] === 'MAINTENANCE' && !$this->ReadPropertyBoolean(self::PROP_SHOWMAINTENANCEENTRIES)){
                unset ($HM_array[$key]);
            }

            if ($HM_dev['HM_devtype'] === 'VIRTUAL_KEY' && !$this->ReadPropertyBoolean(self::PROP_SHOWVIRTUALKEYENTRIES)){
                unset ($HM_array[$key]);
            }
        }



        // Sort device list
        //
        $SortOrder = $this->ReadPropertyInteger('SortOrder');
        switch ($SortOrder) {
            case 3: // by IPS-dev_name
                usort($HM_array, 'self::usort_IPS_dev_name');
                break;
            case 0: // by HM-address
                usort($HM_array, 'self::usort_HM_address');
                break;
            case 1: // by HM-device
                usort($HM_array, 'self::usort_HM_device_adr');
                break;
            case 2: //by HM-type
                usort($HM_array, 'self::usort_HM_devtype');
                break;
            case 4: //by HM-device-name
                usort($HM_array, 'self::usort_HM_devname');
                break;
            default:
                trigger_error('Unknown SortOrder: ' . $SortOrder);
        }

        //print_r($HM_array);

        if ($this->ReadPropertyBoolean(self::PROP_SAVEDEVICELISTINVARIABLE)) {
            //SetValueString($this->GetIDForIdent('DeviceList'), json_encode($HM_array));
            $this->SetValue('DeviceList', json_encode($HM_array)); //array in String variable speichern
        }

        // Generate HTML output code

        $this->UpdateFormField('ProgressBar', 'current', $progressBarCounter++);

        $HTML_intro = "<table width='100%' border='0' align='center' bgcolor=" . self::BG_COLOR_GLOBAL . '>';

        $HTML_ifcs = "<tr style='vertical-align: top'>";
        $HTML_ifcs .= "<td><table style='text-align: left;font-size: large; color: #99AABB'><tr><td><b>HM Inventory ($moduleVersion) </b>";
        $HTML_ifcs .= "<b>&nbsp found at " . strftime('%d.%m.%Y %X', time()) . '</b></td></tr>';
        $HTML_ifcs .= "<tr><td style='font-size: small; color: #CCCCCC'>" . sprintf(
                '%s HomeMatic interfaces (%s connected) with %s HM-RF devices, %s HM-wired devices and %s HmIP devices',
                $HM_interface_num,
                $HM_interface_connected_num,
                $hm_RF_parent_devices_count,
                $hm_Wired_parent_devices_count,
                $hm_IP_parent_devices_count
            ) . '</td>';
        $HTML_ifcs .= "<tr><td style='font-size: small; color: #CCCCCC'>" . sprintf ('%s IPS instances (connected to %s HM channels)', $IPS_device_num, $IPS_HM_channel_num) . '</td>';
        $HTML_ifcs .= '</table></td>';
        $HTML_ifcs .= "<td style='vertical-align: top'>&nbsp;</td>";

        $HTML_ifcs .= "<td style='width: 40%; vertical-align: bottom;'><table style='width: 100%; text-align: right; background-color: " . self::BG_COLOR_INTERFACE_LIST . '\'>';
        //print_r($hm_BidCos_Ifc_list);
        foreach ($hm_BidCos_Ifc_list as $hm_ifce) {
            $dtifc_td_b = "<td style='font-size: small;color: #EEEEEE'>" . ($hm_ifce['DEFAULT'] ? '<i>' : '');
            $dtifc_td_e = ($hm_ifce['DEFAULT'] ? '</i>' : '') . '</td>';
            $dsc_strg   = sprintf('%s', $hm_ifce['CONNECTED'] ? 'connected' : 'Not connected');
            $ifce_info = sprintf('%s (Fw: %s, DC: %s%%)', $hm_ifce['ADDRESS'], $hm_ifce['FIRMWARE_VERSION'], $hm_ifce['DUTY_CYCLE']);
            $HTML_ifcs  .= '<tr>' . $dtifc_td_b . 'Interface: ' . $ifce_info . '&nbsp' . $dtifc_td_e;
            $HTML_ifcs  .= $dtifc_td_b . $hm_ifce['DESCRIPTION'] . $dtifc_td_e . $dtifc_td_b . $dsc_strg . $dtifc_td_e . '</tr>';
        }
        $HTML_ifcs .= '</table></td></tr>';

        $HTML_sep = "<tr><td colspan=3><table style='width: 100%; text-align: left'> <hr><tr><td> </td></tr></table></td></tr>";

        $dthdr_td_b   = "<td style='font-size: small; color: #EEEEEE'><b>";
        $dthdr_td_b_r = "<td style='text-align: right; font-size: small; color: #EEEEEE'><b>";
        $dthdr_td_e   = '</b></td>';
        $dthdr_td_eb  = $dthdr_td_e . $dthdr_td_b;
        $HTML_dvcs    = "<tr><td colspan=3><table style='width: 100%; text-align: left'>";
        $HTML_dvcs    .= '<tr bgcolor=' . self::BG_COLOR_HEADLINE . '>';
        $HTML_dvcs    .= $dthdr_td_b_r . '&nbsp##&nbsp' . $dthdr_td_eb . 'IPS ID' . $dthdr_td_eb . 'IPS device name&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp'
                         . $dthdr_td_eb . 'HM address' . $dthdr_td_e;
        if ($this->ReadPropertyBoolean('ShowHMConfiguratorDeviceNames')) {
            $HTML_dvcs .= $dthdr_td_b . 'HM device name' . $dthdr_td_e;
        }
        $HTML_dvcs .= $dthdr_td_b . 'HM device type' . $dthdr_td_eb . 'Fw.' . $dthdr_td_eb . 'HM channel type' . $dthdr_td_eb . 'Dir.' . $dthdr_td_eb
                      . 'AES' . $dthdr_td_e;
        $HTML_dvcs .= "<td style='width: 2%; text-align: center; color: #EEEEEE; font-size: medium'>Roa- ming</td>";
        foreach ($hm_BidCos_Ifc_list as $hm_ifce) {
            if ($hm_ifce['CONNECTED']) {
                $HTML_dvcs .= "<td style='width: 6%; text-align: center; color: #EEEEEE; font-size: small'>" . $hm_ifce['ADDRESS'] . ' tx/rx&nbsp(db&micro;V)'
                              . '</td>';
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
            $font_tag      = "<font size='2' color=" . (($HM_dev['IPS_HM_d_assgnd'] === false) ? '#DDDDDD' : '#FFAAAA') . '>';
            $dtdvc_td_b    = '<td>' . $font_tag;
            $dtdvc_td_ar_b = "<td style='text-align: right'>" . $font_tag;
            $dtdvc_td_ac_b = "<td style='text-align: center'>" . $font_tag;
            $dtdvc_td_e    = '</font></td>';
            $dtdvc_td_eb   = $dtdvc_td_e . $dtdvc_td_b;
            if (($entry_no++ % 2) === 0) {
                $r_bgcolor = self::BG_COLOR_ODDLINE;
            } else {
                $r_bgcolor = self::BG_COLOR_EVENLINE;
            }
            $HTML_dvcs .= '<tr bgcolor=' . $r_bgcolor . '>' . $dtdvc_td_ar_b . $entry_no . '&nbsp&nbsp' . $dtdvc_td_eb;
            $HTML_dvcs .= $HM_dev['IPS_id'] . $dtdvc_td_eb . utf8_decode($HM_dev['IPS_name']) . $dtdvc_td_eb . $HM_dev['HM_address'] . $dtdvc_td_eb;
            if ($this->ReadPropertyBoolean('ShowHMConfiguratorDeviceNames')) {
                $HTML_dvcs .= utf8_decode($HM_dev['HM_devname']) . $dtdvc_td_eb;
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
                                    $lvl_strg_color = '<font color=#DDDD66>';
                                } else {
                                    $lvl_strg_color = '<font color=#FFFF88>';
                                }
                            } else {
                                $lvl_strg_color = '<font color=#DDDDDD>';
                            }

                            //rx_lvl
                            if ($lciValue[0] !== 65536) {
                                $rx_strg = (string)$lciValue[0];
                            } else {
                                $rx_strg = '--';
                            }
                            //tx_lvl
                            if ($lciValue[1] !== 65536) {
                                $tx_strg = (string)$lciValue[1];
                            } else {
                                $tx_strg = '--';
                            }
                            if (($HM_dev['HM_Roaming'] === '+') || $lciValue[2]) {
                                $lvl_strg = sprintf(
                                    '%s<ins>%s &#047 %s</ins></font>',
                                    $lvl_strg_color,
                                    $rx_strg,
                                    $tx_strg
                                );
                            } else {
                                $lvl_strg = sprintf(
                                    '%s%s &#047 %s</font>',
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
                for ($lci = 0; $lci < $HM_interface_connected_num; $lci++) {
                    $HTML_dvcs .= '<td> </td>';
                }
            }
        }

        if ($HM_module_num === 0) {
            $HTML_dvcs .= "<tr><td colspan=20 style='text-align: center; color: #DDDDDD; font-size: large'><br/>No HomeMatic devices found!</td></tr>";
        }

        $HTML_dvcs .= '</td></tr>';

        // Some comments
        //
        $HTML_notes = "<tr><td colspan=20><table style='width: 100%; text-align: left; color: #666666'><hr><tr><td> </td></tr></table></td></tr>";
        $HTML_notes .= "<tr><td colspan=20><table style='width: 100%; text-align: left; font-size:medium; color: #DDDDDD'><tr><td>Notes:</td></tr>";
        $HTML_notes .= "<tr><td style='font-size: smaller; color: #DDDDDD'><ol>";
        $HTML_notes .= '<li>Interfaces: bold letters indicate the default BidCos-Interface.</li>';
        $HTML_notes .= '<li>Level-pairs: the left value is showing the last signal level received by the device from the interface,';
        $HTML_notes .= ' while the right value is showing the last signal level received by the interface from the device.</li>';
        $HTML_notes .= '<li>Level-pairs: underlined letters of the level-pair indicate the BidCos-Interface associated with the device';
        $HTML_notes .= ' (or all interfaces when Roaming is enabled for the device).</li>';
        $HTML_notes .= '<li>Level-pairs: the yellow level-pair indicates the BidCos-Interface with best signal quality.</li>';
        $HTML_notes .= "<li>Devices without level-pairs haven't send/received anything since last start of the BidCos-service or are wired.</li>";
        $HTML_notes .= '<li>BidCos channels assigned to more than one IPS-device are shown in red.</li>';
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
            fwrite($HTML_file, '<html><head><style>');
            fwrite($HTML_file, 'html,body {font-family:Arial,Helvetica,sans-serif;font-size:12px;background-color:#000000;color:#dddddd;}');
            fwrite($HTML_file, '</style></head><body>');
            fwrite($HTML_file, $HTML_intro);
            fwrite(
                $HTML_file,
                "<tr><td colspan=3><table style='width: 100%; text-align: left; background-color: #112233'><tr><td><h1>HM inventory</h1></td></tr></table></td></tr>"
            );
            fwrite($HTML_file, $HTML_ifcs);
            fwrite($HTML_file, $HTML_sep);
            fwrite($HTML_file, $HTML_dvcs);
            fwrite($HTML_file, $HTML_notes);
            fwrite($HTML_file, $HTML_end);
            fwrite($HTML_file, '</body></html>');
            return fclose($HTML_file);
        }

        return false;
    }

    private function RegisterProperties(): void
    {
        $this->RegisterPropertyBoolean(self::PROP_SAVEDEVICELISTINVARIABLE, false);
        $this->RegisterPropertyBoolean('SaveHMArrayInVariable', false);
        $this->RegisterPropertyString(self::PROP_OUTPUTFILE, IPS_GetKernelDir() . 'HM_inventory.html');
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
        if ($this->HasActiveParent()) {
            $this->SetStatus(IS_ACTIVE);
        } else {
            $this->SetStatus(IS_INACTIVE);
        }
    }

 private function SendRequestMessage(string $methodName, array $params, string $BidCos_Service_adr, $UseSSL, string $Password, string $Username)
    {
     $xml_BidCos_client = new xmlrpc_client($BidCos_Service_adr);
     if ($UseSSL) {
         $xml_BidCos_client->setSSLVerifyHost(0);
         $xml_BidCos_client->setSSLVerifyPeer(false);
     }
     if ($Password !== '') {
         $xml_BidCos_client->setCredentials($Username, $Password);
     }

     $xml_reqmsg = new xmlrpcmsg($methodName, $params);

     $this->SendDebug('send (xmlrpc)', sprintf('send (xmlrpc):%s:%s, params: %s', $BidCos_Service_adr, $methodName, json_encode($params)), 0);
     return $xml_BidCos_client->send($xml_reqmsg);

 }
    private static function usort_HM_address(array $a, array $b)
    {
        $result = strcasecmp($a['HM_address'], $b['HM_address']);

        $a_adr = explode(':', $a['HM_address']);
        $b_adr = explode(':', $b['HM_address']);
        if (count($a_adr) === 2 && count($b_adr) === 2 && strcasecmp($a_adr[0], $b_adr[0]) === 0) {
            $result = (int)$a_adr[1] > $b_adr[1];
        }

        return $result;
    }

    /** @noinspection PhpUnusedPrivateMethodInspection */
    private static function usort_IPS_dev_name(array $a, array $b)
    {
        if (($result = strcasecmp($a['IPS_name'], $b['IPS_name'])) === 0) {
            $result = self::usort_HM_address($a, $b);
        }

        return $result;
    }

    private static function usort_HM_device_adr(array $a, array $b)
    {
        if (($result = strcasecmp($a['HM_device'], $b['HM_device'])) === 0) {
            $result = self::usort_HM_address($a, $b);
        }

        return $result;
    }

    private static function usort_HM_devtype(array $a, array $b)
    {
        if (($result = strcasecmp($a['HM_devtype'], $b['HM_devtype'])) === 0) {
            $result = self::usort_HM_address($a, $b);
        }

        return $result;
    }

    private static function usort_HM_devname(array $a, array $b)
    {
        if (($result = strcasecmp($a['HM_devname'], $b['HM_devname'])) === 0) {
            $result = self::usort_HM_address($a, $b);
        }

        return $result;
    }

    private function GetHMChannelName($HMAddress, $HMDeviceAddress)
    {
        $HMScript = 'Name = (xmlrpc.GetObjectByHSSAddress(interfaces.GetAt(0), "' . $HMDeviceAddress . '")).Name();' . PHP_EOL;

        $ScriptReturn = $this->SendScript($HMAddress, $HMScript);
        $HMChannelName = json_decode($ScriptReturn, true)['Name'];

        if (!is_string($HMChannelName)){ //Wenn der ChannelName auf HM Seite leer ist, dann kommt ein leeres Array zurück
            $HMChannelName = '';
        }

        $this->SendDebug(__FUNCTION__, sprintf('HMAddress: %s, HMDeviceAddress: %s -> %s', $HMAddress, $HMDeviceAddress, $HMChannelName), 0);
        return $HMChannelName;
    }

    private function SendScript($HMAddress, $Script)
    {
        $url = 'Script.exe';

        try {
            $HMScriptResult = $this->LoadHMScript($HMAddress, $url, $Script);
            $xml            = @new SimpleXMLElement(utf8_encode($HMScriptResult), LIBXML_NOBLANKS + LIBXML_NONET);
        } catch (Exception $exc) {
            trigger_error($exc->getMessage());
        }
        if (isset($xml)) {
            unset($xml->exec, $xml->sessionId, $xml->httpUserAgent);
            return json_encode($xml);
        }

        return false;
    }

    private function LoadHMScript($HMAddress, $url, $HMScript)
    {
        if ($HMAddress !== '') {
            $header[] = 'Accept: text/plain,text/xml,application/xml,application/xhtml+xml,text/html';
            $header[] = 'Cache-Control: max-age=0';
            $header[] = 'Connection: close';
            $header[] = 'Accept-Charset: UTF-8';
            $header[] = 'Content-type: text/plain;charset="UTF-8"';

            $ParentId     = @IPS_GetInstance($this->InstanceID)['ConnectionID'];
            $ParentConfig = json_decode(IPS_GetConfiguration($ParentId), true);

            if ($ParentConfig['UseSSL']) {
                $ch = curl_init(sprintf('https://%s:%s/%s', $HMAddress, $ParentConfig['HSSSLPort'], $url));
                curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
                curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, false);
            } else {
                $ch = curl_init(sprintf('http://%s:%s/%s', $HMAddress, $ParentConfig['HSPort'], $url));
            }
            curl_setopt($ch, CURLOPT_HEADER, false);
            curl_setopt($ch, CURLOPT_FAILONERROR, true);
            curl_setopt($ch, CURLOPT_POST, true);
            curl_setopt($ch, CURLOPT_POSTFIELDS, $HMScript);
            curl_setopt($ch, CURLOPT_HTTPHEADER, $header);
            curl_setopt($ch, CURLOPT_HTTPHEADER, ['Expect:']);
            curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
            curl_setopt($ch, CURLOPT_CONNECTTIMEOUT_MS, 1000);
            curl_setopt($ch, CURLOPT_HTTP_VERSION, CURL_HTTP_VERSION_1_1);
            curl_setopt($ch, CURLOPT_TIMEOUT_MS, 5000);

            if ($ParentConfig['Password'] !== '') {
                curl_setopt($ch, CURLOPT_HTTPAUTH, CURLAUTH_BASIC);
                curl_setopt($ch, CURLOPT_USERPWD, $ParentConfig['Username'] . ':' . $ParentConfig['Password']);
            }

            $result = curl_exec($ch);
            $http_code = curl_getinfo($ch, CURLINFO_HTTP_CODE);

            curl_close($ch);

            if (($result === false) || ($http_code >= 400)) {
                trigger_error('CCU unreachable', E_USER_ERROR);
            }
            return $result;
        }

        trigger_error('CCU Address not set.');
        return false;
    }
}
