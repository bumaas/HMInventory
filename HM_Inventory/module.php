<?php
declare(strict_types=1);

// We need the "xmlrpc" include file
// see https://github.com/gggeek/phpxmlrpc/releases
include_once __DIR__ . '/../libs/phpxmlrpc-4.3.0/lib/xmlrpc.inc';

// Klassendefinition

/** @noinspection AutoloadingIssuesInspection */
class HMInventoryReportCreator extends IPSModule
{
    private const STATUS_INST_IP_IS_INVALID = 204; //IP Adresse ist ungültig

    // Some color options for the HTML output
    private const BG_COLOR_GLOBAL         = '#181818';         // Global background color
    private const BG_COLOR_INTERFACE_LIST = '#223344';         // Background color for the interface list
    private const BG_COLOR_HEADLINE       = '#334455';         // Background color for the header line of the device list
    private const BG_COLOR_ODDLINE        = '#181818';         // Background color for the odd lines of the device list
    private const BG_COLOR_EVENLINE       = '#1A2B3C';         // Background color for the even lines of the device list

    private const VERSION = '1.7.2';

    // Überschreibt die interne IPS_Create($id) Funktion
    public function Create()
    {
        // Diese Zeile nicht löschen.
        parent::Create();

        $this->RegisterProperties();

        $this->RegisterTimer('Update', 0, 'HMI_CreateReport(' . $this->InstanceID . ');');
    }

    public function ApplyChanges()
    {
        //Never delete this line!
        parent::ApplyChanges();

        $this->SetTimerInterval('Update', $this->ReadPropertyInteger('UpdateInterval') * 60 * 1000);

        $this->RegisterVariables();

        $this->SetInstanceStatus();
    }

    /**
     * Die folgenden Funktionen stehen automatisch zur Verfügung, wenn das Modul über die "Module Control" eingefügt wurden.
     * Die Funktionen werden, mit dem selbst eingerichteten Prefix, in PHP und JSON-RPC wie folgt zur Verfügung gestellt:.
     */
    public function CreateReport(): void
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
        $IP_adr_BidCos_Service = $this->ReadPropertyString('Host');

        $BidCos_Wired_Service_adr = sprintf('http://%s:2000', $IP_adr_BidCos_Service);
        $BidCos_RF_Service_adr    = sprintf('http://%s:2001', $IP_adr_BidCos_Service);
        $BidCos_IP_Service_adr    = sprintf('http://%s:2010', $IP_adr_BidCos_Service);

        $hm_RF_dev_list    = [];
        $hm_IP_dev_list    = [];
        $hm_Wired_dev_list = [];
        $err               = 0;

        $xml_reqmsg = new xmlrpcmsg('listDevices');

        // get the RF devices
        $xml_BidCos_RF_client = new xmlrpc_client($BidCos_RF_Service_adr);
        $this->SendDebug('send (xmlrpc):', $BidCos_RF_Service_adr . ':listDevices', 0);
        $xml_rtnmsg = $xml_BidCos_RF_client->send($xml_reqmsg);
        if ($xml_rtnmsg->errno === 0) {
            $this->SendDebug('received (xmlrpc):', json_encode($xml_rtnmsg->value()), 0);
            $hm_RF_dev_list = php_xmlrpc_decode($xml_rtnmsg->value());
            //print_r($hm_RF_dev_list);
        } else {
            $this->SendDebug('Error', "Can't get any device information from the BidCos-RF-Service", 0);
            $err++;
        }

        // get the IP devices
        $xml_BidCos_IP_client = new xmlrpc_client($BidCos_IP_Service_adr);
        $xml_reqmsg           = new xmlrpcmsg('listDevices');
        $this->SendDebug('send (xmlrpc):', $BidCos_IP_Service_adr . ':listDevices', 0);
        $xml_rtnmsg = $xml_BidCos_IP_client->send($xml_reqmsg);
        if ($xml_rtnmsg->errno === 0) {
            $this->SendDebug('received (xmlrpc):', json_encode($xml_rtnmsg->value()), 0);
            $hm_IP_dev_list = php_xmlrpc_decode($xml_rtnmsg->value());
            //print_r($hm_IP_dev_list);
        } else {
            $this->SendDebug('Error', "Can't get any device information from the BidCos-IP-Service", 0);
            $err += 2;
        }

        // get the Wired devices
        $xml_BidCos_Wired_client = new xmlrpc_client($BidCos_Wired_Service_adr);
        $xml_reqmsg              = new xmlrpcmsg('listDevices');
        $this->SendDebug('send (xmlrpc):', $BidCos_Wired_Service_adr . ':listDevices', 0);
        $xml_rtnmsg = $xml_BidCos_Wired_client->send($xml_reqmsg);
        if ($xml_rtnmsg->errno === 0) {
            $this->SendDebug('received (xmlrpc):', json_encode($xml_rtnmsg->value()), 0);
            $hm_Wired_dev_list = php_xmlrpc_decode($xml_rtnmsg->value());
            //print_r($hm_Wired_dev_list);
        } else {
            $this->SendDebug('Error', "Can't get any device information from the BidCos-Wired-Service", 0);
            $err += 4;
        }

        // merge all devices
        $hm_dev_list = array_merge($hm_RF_dev_list, $hm_IP_dev_list, $hm_Wired_dev_list);
        if (count($hm_dev_list) === 0) {
            trigger_error("Fatal error: Can't get any device information from the BidCos-Services (Error: $err)");
        }
        //print_r($hm_dev_list);

        // get all Bicos Interfaces
        $xml_reqmsg = new xmlrpcmsg('listBidcosInterfaces');
        $this->SendDebug('send (xmlrpc):', $BidCos_RF_Service_adr . ':listBidcosInterfaces', 0);
        $xml_rtnmsg = $xml_BidCos_RF_client->send($xml_reqmsg);
        if ($xml_rtnmsg->errno === 0) {
            $this->SendDebug('received (xmlrpc):', json_encode($xml_rtnmsg->value()), 0);
            $hm_BidCos_Ifc_list = php_xmlrpc_decode($xml_rtnmsg->value());
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

        $IPS_device_num     = 0;
        $IPS_HM_channel_num = 0;
        $HM_module_num      = 0;
        $HM_array           = [];

        // Fill array with all HM-devices found in IP-Symcon
        //
        foreach (IPS_GetInstanceListByModuleID('{EE4A81C6-5C90-4DB7-AD2F-F6BBD521412E}') as $id) {
            //first check if the device is assigned to the right gateway
            if ($IP_adr_BidCos_Service !== IPS_GetProperty(IPS_GetInstance($id)['ConnectionID'], 'Host')) {
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
                        $HM_devname = $this->GetHMChannelName($IP_adr_BidCos_Service, $hm_dev['ADDRESS']);
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
                unset ($HM_dev);
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
                'HM_Roaming'      => $HM_Roaming];
        }

        // Add HM_devices known by BidCos but not present in IP-Symcon
        //
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
                    if ($hm_dev['TYPE'] === 'VIRTUAL_KEY' && !$this->ReadPropertyBoolean('ShowVirtualKeyEntries')) {
                        continue;
                    }
                    if ($hm_dev['TYPE'] === 'MAINTENANCE' && !$this->ReadPropertyBoolean('ShowMaintenanceEntries')) {
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
                            $HM_devname = $this->GetHMChannelName($IP_adr_BidCos_Service, $HM_address);
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
                            'HM_Roaming'      => $HM_Roaming];
                    }
                }
            }
        }
        // Force communication for RF-level update if requested
        //
        /*
        if ($this->ReadPropertyBoolean('RequestStatusForLevelUpdate')) {
            foreach ($HM_array as &$HM_dev) {
                if (substr($HM_dev['HM_address'], strpos($HM_dev['HM_address'], ':'), 2) == ':1') {
                    $xml_method = new xmlrpcmsg('getParamset', [new xmlrpcval($HM_dev['HM_address'], 'string'), new xmlrpcval('VALUES', 'string')]);
                    $xml_rtnmsg = $xml_BidCos_RF_client->send($xml_method);
                    if ($xml_rtnmsg->errno == 0) {
                        $HM_ParamSet = php_xmlrpc_decode($xml_rtnmsg->value());
                    }
                }
            }
        }
        */

        // Request tx/rx RF-levels from BidCos-RF-Service
        //
        $xml_reqmsg = new xmlrpcmsg('rssiInfo');
        $xml_rtnmsg = $xml_BidCos_RF_client->send($xml_reqmsg);

        $hm_lvl_list = [];
        if ($xml_rtnmsg->errno === 0) {
            $hm_lvl_list = php_xmlrpc_decode($xml_rtnmsg->value());
            //print_r($hm_lvl_list);
        } else {
            echo "Warning: Can't get RF-level information from the BidCos-Service ($BidCos_RF_Service_adr) &nbsp&nbsp&nbsp - &nbsp&nbsp&nbsp ($xml_rtnmsg->errstr)<br>\n";
        }

        // Add tx/rx RF-levels for each device/interface
        //
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
                                    false];
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
                        unset ($hm_lvl);
                    }
                    $HM_dev['HM_levels'] = $HM_lvl_array;
                }
            }
            unset ($HM_dev);
        }

        // Request tx/rx RF-levels from BidCos-IP-Service
        //
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
                $xml_method = new xmlrpcmsg(
                    'getParamset', [
                                     new xmlrpcval($hm_adr[0] . ':0', 'string'),
                                     new xmlrpcval('VALUES', 'string')]
                );
                $xml_rtnmsg = $xml_BidCos_IP_client->send($xml_method);
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
        unset ($HM_dev);

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

        if ($this->ReadPropertyBoolean('SaveDeviceListInVariable')) {
            //SetValueString($this->GetIDForIdent('DeviceList'), json_encode($HM_array));
            $this->SetValue('DeviceList', json_encode($HM_array)); //array in String variable speichern
        }

        // Generate HTML output code

        $HTML_intro = "<table width='100%' border='0' align='center' bgcolor=" . self::BG_COLOR_GLOBAL . '>';

        //$HTML_ifcs = "<tr valign='top' width='100%'>";
        $HTML_ifcs = "<tr valign='top'>";
        $HTML_ifcs .= "<td><table align='left'><tr><td><font size='3' color='#99AABB'><b>HM Inventory (" . self::VERSION . ') </font></b>';
        $HTML_ifcs .= "<font size='3' color='#CCCCCC'><b>&nbsp found at " . strftime('%d.%m.%Y %X', time()) . '</font></b></td></tr>';
        $HTML_ifcs .= "<tr><td><font size='2' color='#CCCCCC'>" . $HM_interface_num . ' HomeMatic interfaces (' . $HM_interface_connected_num
                      . ' connected)</td>';
        $HTML_ifcs .= "<tr><td><font size='2' color='#CCCCCC'>" . $IPS_device_num . ' IPS instances (connected to ' . $IPS_HM_channel_num
                      . ' HM channels)</td>';
        $HTML_ifcs .= '</table></td>';
        $HTML_ifcs .= "<td valign='top'>&nbsp;</td>";

        $HTML_ifcs .= "<td width='40%' valign='bottom'><table width='100%' align='right' bgcolor=" . self::BG_COLOR_INTERFACE_LIST . '>';
        //print_r($hm_BidCos_Ifc_list);
        foreach ($hm_BidCos_Ifc_list as $hm_ifce) {
            $dtifc_td_b = "<td><font size='2' color='#EEEEEE'>" . ($hm_ifce['DEFAULT'] ? '<i>' : '');
            $dtifc_td_e = ($hm_ifce['DEFAULT'] ? '</i>' : '') . '</font></td>';
            $dsc_strg   = sprintf('%sconnected', $hm_ifce['CONNECTED'] ? '' : 'Not ');
            $HTML_ifcs  .= '<tr>' . $dtifc_td_b . 'Interface: ' . $hm_ifce['ADDRESS'] . '&nbsp' . $dtifc_td_e;
            $HTML_ifcs  .= $dtifc_td_b . $hm_ifce['DESCRIPTION'] . $dtifc_td_e . $dtifc_td_b . $dsc_strg . $dtifc_td_e . '</tr>';
        }
        $HTML_ifcs .= '</table></td></tr>';

        $HTML_sep = "<tr><td colspan=3><table width='100%' align='left'> <hr><tr><td> </td></tr></table></td></tr>";

        $dthdr_td_b   = "<td><font size='2' color='#EEEEEE'><b>";
        $dthdr_td_b_r = "<td align='right'><font size='2' color='#EEEEEE'><b>";
        $dthdr_td_e   = '</font></b></td>';
        $dthdr_td_eb  = $dthdr_td_e . $dthdr_td_b;
        $HTML_dvcs    = "<tr><td colspan=3><table width='100%' align='left'>";
        $HTML_dvcs    .= '<tr bgcolor=' . self::BG_COLOR_HEADLINE . '>';
        $HTML_dvcs    .= $dthdr_td_b_r . '&nbsp##&nbsp' . $dthdr_td_eb . 'IPS ID' . $dthdr_td_eb . 'IPS device name&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp'
                         . $dthdr_td_eb . 'HM address' . $dthdr_td_e;
        if ($this->ReadPropertyBoolean('ShowHMConfiguratorDeviceNames')) {
            $HTML_dvcs .= $dthdr_td_b . 'HM device name' . $dthdr_td_e;
        }
        $HTML_dvcs .= $dthdr_td_b . 'HM device type' . $dthdr_td_eb . 'Fw.' . $dthdr_td_eb . 'HM channel type' . $dthdr_td_eb . 'Dir.' . $dthdr_td_eb
                      . 'AES' . $dthdr_td_e;
        $HTML_dvcs .= "<td width='2%' align='center'><font size='2' color='#EEEEEE'> Roa- ming" . '</font></td>';
        foreach ($hm_BidCos_Ifc_list as $hm_ifce) {
            if ($hm_ifce['CONNECTED']) {
                $HTML_dvcs .= "<td width='6%' align='center'><font size='2' color='#EEEEEE'>" . $hm_ifce['ADDRESS'] . ' tx/rx&nbsp(db&micro;V)'
                              . '</font></td>';
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
            $dtdvc_td_ar_b = "<td align='right'>" . $font_tag;
            $dtdvc_td_ac_b = "<td align='center'>" . $font_tag;
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

                            if (($HM_dev['HM_Roaming'] === '+') || $lciValue[2]) {
                                $fmt_strg = '%s<ins>%s &#047 %s</ins></font>';
                            } else {
                                $fmt_strg = '%s%s &#047 %s</font>';
                            }

                            //rx_lvl
                            if ($lciValue[0] !== 65536) {
                                $rx_strg = (string) $lciValue[0];
                            } else {
                                $rx_strg = '--';
                            }
                            //tx_lvl
                            if ($lciValue[1] !== 65536) {
                                $tx_strg = (string) $lciValue[1];
                            } else {
                                $tx_strg = '--';
                            }
                            $lvl_strg = sprintf(
                                $fmt_strg, $lvl_strg_color, $rx_strg, $tx_strg
                            );

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
            $HTML_dvcs .= "<tr><td colspan=20 align='center'><br/><font size='4' color='#DDDDDD'>No HomeMatic devices found!</font></td></tr>";
        }

        $HTML_dvcs .= '</td></tr>';

        // Some comments
        //
        $HTML_notes = "<tr><td colspan=20><table width='100%' align='left'><hr><color='#666666'><tr><td> </td></tr></table></td></tr>";
        $HTML_notes .= "<tr><td colspan=20><table width='100%' align='left'><tr><td><font size='3' color='#DDDDDD'>Notes:</font></td></tr>";
        $HTML_notes .= "<tr><td><font size='2' color='#DDDDDD'><ol>";
        $HTML_notes .= '<li>Interfaces: bold letters indicate the default BidCos-Interface.</li>';
        $HTML_notes .= '<li>Level-pairs: the left value is showing the last signal level received by the device from the interface,';
        $HTML_notes .= ' while the right value is showing the last signal level received by the interface from the device.</li>';
        //$HTML_notes .= "<li>Level-pairs: italic letters of the level-pair indicate the BidCos-Interface associated with the device";
        $HTML_notes .= '<li>Level-pairs: underlined letters of the level-pair indicate the BidCos-Interface associated with the device';
        $HTML_notes .= ' (or all interfaces when Roaming is enabled for the device).</li>';
        $HTML_notes .= '<li>Level-pairs: the yellow level-pair indicates the BidCos-Interface with best signal quality.</li>';
        $HTML_notes .= "<li>Devices without level-pairs haven't send/received anything since last start of the BidCos-service or are wired.</li>";
        $HTML_notes .= '<li>BidCos channels assigned to more than one IPS-device are shown in red.</li>';
        $HTML_notes .= '</ol></font></td></tr>';
        $HTML_notes .= '</table></td></tr>';

        $HTML_end = '</table>';

        // Output the results

        $OutputFileName = $this->ReadPropertyString('OutputFile');
        if ($OutputFileName) {
            $HTML_file = fopen($OutputFileName, 'wb');
            fwrite($HTML_file, '<html><head><style type="text/css">');
            fwrite($HTML_file, 'html,body {font-family:Arial,Helvetica,sans-serif;font-size:12px;background-color:#000000;color:#dddddd;}');
            fwrite($HTML_file, '</style></head><body>');
            fwrite($HTML_file, $HTML_intro);
            fwrite(
                $HTML_file,
                "<tr><td colspan=3><table width='100%' align='left' bgcolor=#112233><tr><td><h1>HM inventory</h1></td></tr></table></td></tr>"
            );
            fwrite($HTML_file, $HTML_ifcs);
            fwrite($HTML_file, $HTML_sep);
            fwrite($HTML_file, $HTML_dvcs);
            fwrite($HTML_file, $HTML_notes);
            fwrite($HTML_file, $HTML_end);
            fwrite($HTML_file, '</body></html>');
            fclose($HTML_file);
        }
    }

    private function RegisterProperties(): void
    {
        $host = '';
        if (IPS_GetKernelRunlevel() === KR_READY) { //Kernel ready
            //set default of host to host of first HM instance
            $HMinstanceList = IPS_GetInstanceListByModuleID('{A151ECE9-D733-4FB9-AA15-7F7DD10C58AF}');
            if (count($HMinstanceList) > 0) {
                $host = IPS_GetProperty($HMinstanceList[0], 'Host');
            }
        }

        $this->RegisterPropertyString('Host', $host);
        $this->RegisterPropertyBoolean('SaveDeviceListInVariable', false);
        $this->RegisterPropertyBoolean('SaveHMArrayInVariable', false);
        $this->RegisterPropertyString('OutputFile', IPS_GetKernelDir() . 'HM_inventory.html');
        $this->RegisterPropertyInteger('SortOrder', 0);
        $this->RegisterPropertyBoolean('ShowVirtualKeyEntries', false);
        $this->RegisterPropertyBoolean('ShowMaintenanceEntries', true);
        $this->RegisterPropertyBoolean('ShowNotUsedChannels', true);
        $this->RegisterPropertyBoolean('ShowLongIPSDeviceNames', false);
        $this->RegisterPropertyBoolean('ShowHMConfiguratorDeviceNames', true);
        $this->RegisterPropertyInteger('UpdateInterval', 0);
    }

    private function RegisterVariables(): void
    {
        if ($this->ReadPropertyBoolean('SaveDeviceListInVariable')) {
            $this->RegisterVariableString('DeviceList', 'Device Liste', '', 1);
        }
    }

    private function SetInstanceStatus(): bool
    {
        //IP Prüfen
        $ip = $this->ReadPropertyString('Host');
        if (filter_var($ip, FILTER_VALIDATE_IP)) {
            $this->SetStatus(IS_ACTIVE);
        } else {
            $this->SetStatus(self::STATUS_INST_IP_IS_INVALID); //IP Adresse ist ungültig
        }

        return true;
    }

    private static function usort_HM_address(array $a, array $b)
    {
        $result = strcasecmp($a['HM_address'], $b['HM_address']);

        $a_adr = explode(':', $a['HM_address']);
        $b_adr = explode(':', $b['HM_address']);
        if (count($a_adr) === 2 && count($b_adr) === 2 && strcasecmp($a_adr[0], $b_adr[0]) === 0) {
            $result = (int) $a_adr[1] > $b_adr[1];
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

    /** @noinspection PhpUnusedPrivateMethodInspection */
    private static function usort_HM_device_adr(array $a, array $b)
    {
        if (($result = strcasecmp($a['HM_device'], $b['HM_device'])) === 0) {
            $result = self::usort_HM_address($a, $b);
        }

        return $result;
    }

    /** @noinspection PhpUnusedPrivateMethodInspection */
    private static function usort_HM_devtype(array $a, array $b)
    {
        if (($result = strcasecmp($a['HM_devtype'], $b['HM_devtype'])) === 0) {
            $result = self::usort_HM_address($a, $b);
        }

        return $result;
    }

    /** @noinspection PhpUnusedPrivateMethodInspection */
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
        return json_decode($this->SendScript($HMAddress, $HMScript), true)['Name'];
    }

    private function SendScript($HMAddress, $Script)
    {
        $url = 'Script.exe';

        try {
            $HMScriptResult = $this->LoadHMScript($HMAddress, $url, $Script);
            $xml            = @new SimpleXMLElement(utf8_encode($HMScriptResult), LIBXML_NOBLANKS + LIBXML_NONET);
        }
        catch(Exception $exc) {
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

            $this->SendDebug('send (curl):', $HMScript, 0);

            $ch = curl_init('http://' . $HMAddress . ':8181/' . $url);
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
            $result = curl_exec($ch);
            $this->SendDebug('received (curl):', $result, 0);

            $http_code = curl_getinfo($ch, CURLINFO_HTTP_CODE);

            curl_close($ch);
            if ($http_code >= 400) {
                trigger_error('CCU unreachable:' . $http_code);
            }
            if ($result === false) {
                trigger_error('CCU unreachable');
            }
            return $result;
        }

        trigger_error('CCU Address not set.');
        return false;
    }
}
