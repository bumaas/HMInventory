<?php /** @noinspection CurlSslServerSpoofingInspection */

declare(strict_types=1);

use PhpXmlRpc\Autoloader;
use PhpXmlRpc\Client;
use PhpXmlRpc\Encoder;
use PhpXmlRpc\Request;
use PhpXmlRpc\Response;

// see https://github.com/gggeek/phpxmlrpc (vendored: libs/phpxmlrpc, nur src/)
require_once __DIR__ . '/../libs/phpxmlrpc/src/Autoloader.php';
Autoloader::register();

/** @noinspection AutoloadingIssuesInspection */
class HMInventoryReportCreator extends IPSModuleStrict
{
    private const array  SERVICETYPES = ['RF', 'IP', 'WR'];
    private const string ERROR_MSG    = "Can't get any device information from the BidCoS-%s-Service";

    private const string GUID_HOMEMATIC_DEVICE = '{EE4A81C6-5C90-4DB7-AD2F-F6BBD521412E}';

    // Some color options for the HTML output
    private const string BG_COLOR_INTERFACE_LIST = '#223344';         // Background color for the interface list
    private const int    INVALID_LEVEL           = 65536;

    private const int STATUS_OUTPUTFILE_NOT_CREATABLE = 200;

    private const int PROGRESS_BAR_MAX = 9;

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

    private const string VAR_IDENT_DEVICELIST = 'DeviceList';

    private int $progressBarStep = 0;

    public function Create(): void
    {
        // Diese Zeile nicht löschen.
        parent::Create();

        $this->RegisterProperties();

        $this->RegisterTimer('Update', 0, 'IPS_RequestAction(' . $this->InstanceID . ', "CreateReport", true);');

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

        if (!$this->isOutputFileCreatable($this->ReadPropertyString(self::PROP_OUTPUTFILE))) {
            $this->SetStatus(self::STATUS_OUTPUTFILE_NOT_CREATABLE);
        } else {
            $this->SetInstanceStatus();
        }

        $this->SetSummary($this->ReadPropertyString(self::PROP_OUTPUTFILE));

        $this->updateReportLink();
    }

    public function MessageSink(int $TimeStamp, int $SenderID, int $Message, array $Data): void
    {
        parent::MessageSink($TimeStamp, $SenderID, $Message, $Data);

        if (($Message === IPS_KERNELMESSAGE) && ($Data[0] === KR_READY)) {
            $this->ApplyChanges();
        }
    }

    public function GetConfigurationForm(): string
    {
        $this->RegisterOnceTimer('UpdateReportLink', 'IPS_RequestAction(' . $this->InstanceID . ', "UpdateReportLink", 0);');

        $form = file_get_contents(__DIR__ . DIRECTORY_SEPARATOR . 'form.json');
        return $form !== false ? $form : '';
    }

    public function RequestAction(string $Ident, mixed $Value): void
    {
        switch ($Ident) {
            case 'UpdateReportLink':
                $this->updateReportLink();
                break;
            case 'CreateReport':
                $this->CreateReport();
                break;
            default:
                throw new InvalidArgumentException('Invalid ident: ' . $Ident);
        }
    }

    public function ReceiveData(string $JSONString): string
    {
        $this->LogMessage(sprintf('Fatal error: no ReceiveData expected. (%s)', $JSONString), KL_ERROR);

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
     */
    public function CreateReport(): bool
    {
        $this->progressBarInit();
        $this->advanceProgressBar(); // Schritt 1

        $reportData = $this->collectDeviceData();
        if ($reportData === null) {
            $this->UpdateFormField('ProgressBar', 'visible', false);
            return false;
        }

        // Wenn gewünscht, die sortierte Liste in die Variable schreiben
        if ($this->ReadPropertyBoolean(self::PROP_SAVEDEVICELISTINVARIABLE)) {
            $this->SetValue(self::VAR_IDENT_DEVICELIST, json_encode($reportData['HM_array'], JSON_THROW_ON_ERROR));
        }

        // Generate HTML output code
        $this->advanceProgressBar(); // Schritt 6

        $headerHtml     = $this->renderHeaderSection($reportData);
        $interfacesHtml = $this->renderInterfacesSection($reportData['hm_BidCos_Ifc_list']);
        $devicesHtml    = $this->renderDevicesSection($reportData);

        $this->advanceProgressBar(); // Schritt 7

        $notesHtml = $this->renderNotesSection();

        // Output the results
        $outputFileName = $this->ReadPropertyString(self::PROP_OUTPUTFILE);
        if ($outputFileName === '') {
            $this->UpdateFormField('ProgressBar', 'visible', false);
            return false;
        }

        $this->advanceProgressBar(); // Schritt 8
        $resolvedOutputFileName = $this->resolveOutputFilePath($outputFileName);

        $htmlContent = $this->getHtmlContent($headerHtml, $interfacesHtml, $this->renderSeparator(), $devicesHtml, $notesHtml);

        if (@file_put_contents($resolvedOutputFileName, $htmlContent) === false) {
            echo sprintf($this->Translate('File "%s" not writable!') . PHP_EOL, $resolvedOutputFileName);
            return false;
        }

        $this->UpdateFormField('ProgressBar', 'current', self::PROGRESS_BAR_MAX); // Finaler Schritt 9
        IPS_Sleep(200);
        $this->UpdateFormField('ProgressBar', 'visible', false);
        $this->updateReportLink();

        return true;
    }

    public function GetReportUrl(): string
    {
        $outputFile = $this->ReadPropertyString(self::PROP_OUTPUTFILE);
        return $this->buildUserReportLink($outputFile);
    }

    public function GetOutputFileAbsolutePath(): string
    {
        $outputFile = $this->ReadPropertyString(self::PROP_OUTPUTFILE);
        return $this->resolveOutputFilePath($outputFile);
    }

    //
    // Pfad- und Link-Behandlung
    //

    private function isOutputFileCreatable(string $outputFile): bool
    {
        if ($outputFile === '') {
            return false;
        }

        $resolvedPath = $this->resolveOutputFilePath($outputFile);
        if ($resolvedPath === '') {
            return false;
        }

        $directory = dirname($resolvedPath);
        if (!is_dir($directory) || !is_writable($directory)) {
            return false;
        }

        if (file_exists($resolvedPath) && !is_writable($resolvedPath)) {
            return false;
        }

        return true;
    }

    private function isAbsolutePath(string $path): bool
    {
        if (str_starts_with($path, '/')) {
            return true;
        }

        if (preg_match('/^[A-Za-z]:[\/\\\\]/', $path) === 1) {
            return true;
        }

        if (str_starts_with($path, '\\\\') || str_starts_with($path, '//')) {
            return true;
        }

        return false;
    }

    private function resolveOutputFilePath(string $outputFile): string
    {
        if ($outputFile === '') {
            return '';
        }

        if ($this->isAbsolutePath($outputFile)) {
            return $outputFile;
        }

        return IPS_GetKernelDir() . $outputFile;
    }

    private function updateReportLink(): void
    {
        $outputFile   = $this->ReadPropertyString(self::PROP_OUTPUTFILE);
        $link         = $this->GetReportUrl();
        $resolvedPath = $this->resolveOutputFilePath($outputFile);

        $isVisible = ($link !== '') && ($outputFile !== '') && file_exists($resolvedPath);
        $this->UpdateFormField('ReportLink', 'visible', $isVisible);
    }

    private function buildUserReportLink(string $outputFile): string
    {
        if ($outputFile === '') {
            return '';
        }
        $resolvedOutputFile = $this->resolveOutputFilePath($outputFile);
        if ($resolvedOutputFile === '') {
            return '';
        }

        $kernelDir = IPS_GetKernelDir();
        $userDir   = $kernelDir . 'user' . DIRECTORY_SEPARATOR;

        $normalizedOutput = str_replace(['/', '\\'], DIRECTORY_SEPARATOR, $resolvedOutputFile);
        if (stripos($normalizedOutput, $userDir) !== 0) {
            return '';
        }

        $relativePath = substr($normalizedOutput, strlen($userDir));
        $relativePath = str_replace(DIRECTORY_SEPARATOR, '/', $relativePath);

        $parts       = array_map('rawurlencode', explode('/', $relativePath));
        $relativeUrl = implode('/', $parts);

        $networkInfo = Sys_GetNetworkInfo();
        $serverIp    = 'localhost';
        foreach ($networkInfo as $adapter) {
            if (!empty($adapter['IP'])) {
                $serverIp = $adapter['IP'];
                break;
            }
        }

        $serverPort = '82';

        return 'http://' . $serverIp . ':' . $serverPort . '/user/' . $relativeUrl;
    }

    //
    // Datensammlung
    //

    /**
     * Sammelt alle benötigten Daten für den Report.
     *
     * @return array|null
     * @throws \JsonException
     */
    private function collectDeviceData(): ?array
    {
        $parentId = $this->fetchParentId();
        if (!$this->isGatewayActive($parentId)) {
            return null;
        }

        $parentConfig     = json_decode(IPS_GetConfiguration($parentId), true, 512, JSON_THROW_ON_ERROR);
        $ccuHost          = IPS_GetProperty($parentId, 'Host');
        $rfServiceAddress = $this->getBidCosServiceAddress($ccuHost, $parentConfig, 'RF');

        [$devList, $devCounter, $errorCount] = $this->getDeviceLists($parentId, $parentConfig);
        if (count($devList) === 0) {
            $this->LogMessage("Can't get any device information from the BidCos-Services (Error: $errorCount)", KL_ERROR);
            return null;
        }

        // 1. BidCos-Interfaces
        $this->advanceProgressBar(); // Schritt 2
        $response = $this->SendRequestMessage(
            'listBidcosInterfaces',
            [],
            $rfServiceAddress,
            $parentConfig['UseSSL'],
            $parentConfig['Password'],
            $parentConfig['Username']
        );

        if ($response->errno !== 0) {
            $this->SendDebug('Error', "Can't get HM-interface information from the BidCos-RF-Service", 0);
            return null;
        }

        $ifcList = (new Encoder())->decode($response->value());
        $default = array_column($ifcList, 'DEFAULT');
        array_multisort($default, SORT_DESC, $ifcList);

        $interfaceNum          = count($ifcList);
        $interfaceConnectedNum = 0;
        $defaultInterfaceNo    = 0;
        foreach ($ifcList as $key => $ifc) {
            if ($ifc['CONNECTED']) {
                $interfaceConnectedNum++;
            }
            if ($ifc['DEFAULT']) {
                $defaultInterfaceNo = $key;
            }
        }

        // 2. Geräte, die in IP-Symcon angelegt sind
        $this->advanceProgressBar(); // Schritt 3
        [$deviceRecords, $ipsDeviceNum, $ipsChannelNum, $moduleNum] = $this->collectIpsDevices($parentId, $ccuHost, $devList);

        // 3. Kanäle, die nicht in IP-Symcon genutzt werden (optional)
        $this->advanceProgressBar(); // Schritt 4
        if ($this->ReadPropertyBoolean(self::PROP_SHOWNOTUSEDCHANNELS)) {
            $moduleNum = $this->appendUnusedChannels($deviceRecords, $ccuHost, $devList, $moduleNum);
        }

        // 4. RSSI-Levels
        $this->advanceProgressBar(); // Schritt 5
        $this->attachRssiLevels($deviceRecords, $ifcList, $rfServiceAddress, $parentConfig);

        $this->sortDeviceRecords($deviceRecords);

        return [
            'dev_counter'                => $devCounter,
            'HM_array'                   => $deviceRecords,
            'hm_BidCos_Ifc_list'         => $ifcList,
            'HM_interface_num'           => $interfaceNum,
            'HM_interface_connected_num' => $interfaceConnectedNum,
            'IPS_device_num'             => $ipsDeviceNum,
            'IPS_HM_channel_num'         => $ipsChannelNum,
            'HM_module_num'              => $moduleNum,
            'HM_default_interface_no'    => $defaultInterfaceNo
        ];
    }

    /**
     * Sammelt alle HomeMatic-Geräteinstanzen, die in IP-Symcon angelegt sind.
     *
     * @return array{0: array, 1: int, 2: int, 3: int} [Geräteliste, Anzahl IPS-Instanzen, Anzahl belegter HM-Kanäle, Anzahl Einträge]
     */
    private function collectIpsDevices(int $parentId, string $ccuHost, array $devList): array
    {
        $records       = [];
        $ipsDeviceNum  = 0;
        $ipsChannelNum = 0;
        $moduleNum     = 0;

        foreach (IPS_GetInstanceListByModuleID(self::GUID_HOMEMATIC_DEVICE) as $id) {
            if ($parentId !== IPS_GetInstance($id)['ConnectionID']) {
                continue;
            }
            $moduleNum++;
            $ipsDeviceNum++;

            $address   = IPS_GetProperty($id, 'Address');
            $needlePos = strpos($address, ':');
            if (!$needlePos) {
                continue;
            }
            $parentAddress = substr($address, 0, $needlePos);

            $channelInfo = $this->extractChannelInfo(
                $this->findDeviceByAddress($devList, $address),
                $this->findDeviceByAddress($devList, $parentAddress),
                $ccuHost
            );

            // Ist der HM-Kanal bereits einer anderen IPS-Instanz zugeordnet?
            $alreadyAssigned = false;
            foreach ($records as &$record) {
                if ($record['HM_address'] === $address) {
                    $record['IPS_HM_d_assgnd'] = true;
                    $alreadyAssigned           = true;
                    break;
                }
            }
            unset($record);
            if (!$alreadyAssigned) {
                $ipsChannelNum++;
            }

            $ipsName = $this->ReadPropertyBoolean(self::PROP_SHOWLONGIPSDEVICENAMES) ? IPS_GetLocation($id) : IPS_GetName($id);

            $records[] = $this->createDeviceRecord($moduleNum, $id, $ipsName, $alreadyAssigned, $address, $channelInfo);
        }

        return [$records, $ipsDeviceNum, $ipsChannelNum, $moduleNum];
    }

    /**
     * Ergänzt die Geräteliste um Kanäle, die nicht in IP-Symcon angelegt sind.
     *
     * @return int Fortgeschriebene Anzahl der Einträge
     */
    private function appendUnusedChannels(array &$records, string $ccuHost, array $devList, int $moduleNum): int
    {
        foreach ($devList as $dev) {
            $address = $dev['ADDRESS'];

            if (!isset($dev['PARENT']) || ($dev['PARENT'] === '')) {
                continue;
            }
            if ($this->recordWithAddressExists($records, $address)) {
                continue;
            }
            if (($dev['TYPE'] === 'VIRTUAL_KEY') && !$this->ReadPropertyBoolean(self::PROP_SHOWVIRTUALKEYENTRIES)) {
                continue;
            }
            if (($dev['TYPE'] === 'MAINTENANCE') && !$this->ReadPropertyBoolean(self::PROP_SHOWMAINTENANCEENTRIES)) {
                continue;
            }

            $moduleNum++;
            $parentAddress = substr($address, 0, (int)strpos($address, ':'));
            $channelInfo   = $this->extractChannelInfo($dev, $this->findDeviceByAddress($devList, $parentAddress), $ccuHost);

            $records[] = $this->createDeviceRecord($moduleNum, '-', '-', false, $address, $channelInfo);
        }

        return $moduleNum;
    }

    private function findDeviceByAddress(array $devList, string $address): ?array
    {
        foreach ($devList as $dev) {
            if ($dev['ADDRESS'] === $address) {
                return $dev;
            }
        }
        return null;
    }

    private function recordWithAddressExists(array $records, string $address): bool
    {
        foreach ($records as $record) {
            if ($record['HM_address'] === $address) {
                return true;
            }
        }
        return false;
    }

    /**
     * Extrahiert die Anzeige-Informationen eines HM-Kanals (samt Informationen des zugehörigen Geräts).
     */
    private function extractChannelInfo(?array $channelDev, ?array $parentDev, string $ccuHost): array
    {
        $info = [
            'HM_device'     => '-',
            'HM_devname'    => '-',
            'HM_FWversion'  => ' ',
            'HM_devtype'    => '-',
            'HM_direction'  => '-',
            'HM_AES_active' => '-',
            'HM_Interface'  => '',
            'HM_Roaming'    => ' ',
        ];

        if ($channelDev === null) {
            return $info;
        }

        if (isset($channelDev['PARENT_TYPE'])) {
            $info['HM_device'] = $channelDev['PARENT_TYPE'];
        }

        if ($this->ReadPropertyBoolean(self::PROP_SHOWHMCONFIGURATORDEVICENAMES)) {
            $info['HM_devname'] = $this->getHMChannelName($ccuHost, $channelDev['ADDRESS']);
        }

        $parentInfo           = $this->extractDeviceInfo($parentDev);
        $info['HM_FWversion'] = $parentInfo['FIRMWARE'];
        $info['HM_Interface'] = $parentInfo['INTERFACE'];
        $info['HM_Roaming']   = $parentInfo['ROAMING'];

        $info['HM_devtype']   = $channelDev['TYPE'];
        $info['HM_direction'] = match ($channelDev['DIRECTION'] ?? null) {
            1       => 'TX',
            2       => 'RX',
            default => '-'
        };

        if (($channelDev['AES_ACTIVE'] ?? 0) !== 0) {
            $info['HM_AES_active'] = '+';
        }

        return $info;
    }

    private function extractDeviceInfo(?array $parentDev): array
    {
        if ($parentDev !== null) {
            $firmware  = $parentDev['FIRMWARE'];
            $interface = $parentDev['INTERFACE'] ?? '';
            $roaming   = (isset($parentDev['ROAMING']) && $parentDev['ROAMING']) ? '+' : '-';
        } else {
            $firmware  = '';
            $interface = '';
            $roaming   = '';
        }
        return ['FIRMWARE' => $firmware, 'INTERFACE' => $interface, 'ROAMING' => $roaming];
    }

    private function createDeviceRecord(
        int $occurrence,
        int|string $ipsId,
        string $ipsName,
        bool $alreadyAssigned,
        string $address,
        array $channelInfo
    ): array {
        return [
            'IPS_occ'         => $occurrence,
            'IPS_id'          => $ipsId,
            'IPS_name'        => $ipsName,
            'IPS_HM_d_assgnd' => $alreadyAssigned,
            'HM_address'      => $address,
            'HM_device'       => $channelInfo['HM_device'],
            'HM_devname'      => $channelInfo['HM_devname'],
            'HM_FWversion'    => $channelInfo['HM_FWversion'],
            'HM_devtype'      => $channelInfo['HM_devtype'],
            'HM_direction'    => $channelInfo['HM_direction'],
            'HM_AES_active'   => $channelInfo['HM_AES_active'],
            'HM_Interface'    => $channelInfo['HM_Interface'],
            'HM_Roaming'      => $channelInfo['HM_Roaming']
        ];
    }

    /**
     * Fragt die RSSI-Werte vom BidCos-RF-Service ab und ergänzt sie in der Geräteliste ('HM_levels').
     *
     * @throws \JsonException
     */
    private function attachRssiLevels(array &$records, array $ifcList, string $rfServiceAddress, array $parentConfig): void
    {
        $response = $this->SendRequestMessage(
            'rssiInfo',
            [],
            $rfServiceAddress,
            $parentConfig['UseSSL'],
            $parentConfig['Password'],
            $parentConfig['Username']
        );
        $rssiList = ($response->errno === 0) ? (new Encoder())->decode($response->value()) : [];

        foreach ($records as &$record) {
            // Nur für das Hauptgerät (Kanal 0 oder ohne Kanal) RSSI suchen
            $deviceAddress = explode(':', $record['HM_address'])[0];
            if (!isset($rssiList[$deviceAddress])) {
                continue;
            }

            $record['HM_levels'] = [];
            foreach ($ifcList as $ifc) {
                if (!$ifc['CONNECTED']) {
                    continue;
                }

                $ifcAddress = $ifc['ADDRESS'];
                if (isset($rssiList[$deviceAddress][$ifcAddress])) {
                    // Speichere RX, TX, IsAssociated, IsBest
                    $record['HM_levels'][] = [
                        $rssiList[$deviceAddress][$ifcAddress][0],  // RX
                        $rssiList[$deviceAddress][$ifcAddress][1],  // TX
                        ($record['HM_Interface'] === $ifcAddress),  // IsAssociated
                        false                                       // IsBest (wird unten berechnet)
                    ];
                }
            }

            // Bestes Interface markieren
            if (count($record['HM_levels']) > 0) {
                $bestIdx = 0;
                $maxRssi = -255;
                foreach ($record['HM_levels'] as $idx => $level) {
                    if (($level[0] > $maxRssi) && ($level[0] !== self::INVALID_LEVEL)) {
                        $maxRssi = $level[0];
                        $bestIdx = $idx;
                    }
                }
                $record['HM_levels'][$bestIdx][3] = true;
            }
        }
        unset($record);
    }

    //
    // HTML-Rendering
    //

    private function renderHeaderSection(array $data): string
    {
        $moduleVersion = $this->getModuleVersion();
        $inventoryStr  = sprintf('<b>HM Inventory (%s) </b><b>&nbsp; found at %s</b>', $moduleVersion, date('d.m.Y H:i:s'));
        $interfaceStr  = sprintf(
            '%s HomeMatic interfaces (%s connected) with %s HM-RF devices, %s HM-wired devices and %s HmIP devices',
            $data['HM_interface_num'],
            $data['HM_interface_connected_num'],
            $data['dev_counter']['RF'],
            $data['dev_counter']['WR'],
            $data['dev_counter']['IP']
        );
        $instanceStr   = sprintf('%s IPS instances (connected to %s HM channels)', $data['IPS_device_num'], $data['IPS_HM_channel_num']);

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
        $dthdr_td_b   = '<td style="font-size: small; color: #EEEEEE"><b>';
        $dthdr_td_b_r = '<td style="text-align: right; font-size: small; color: #EEEEEE"><b>';
        $dthdr_td_e   = '</b></td>';
        $dthdr_td_eb  = $dthdr_td_e . $dthdr_td_b;

        $html .= $dthdr_td_b_r . '&nbsp;##&nbsp;' . $dthdr_td_eb . 'IPS ID' . $dthdr_td_eb
                 . 'IPS device name&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' . $dthdr_td_eb . 'HM address' . $dthdr_td_e;

        if ($this->ReadPropertyBoolean(self::PROP_SHOWHMCONFIGURATORDEVICENAMES)) {
            $html .= $dthdr_td_b . 'HM device name' . $dthdr_td_e;
        }

        $html .= $dthdr_td_b . 'HM device type' . $dthdr_td_eb . 'Fw.' . $dthdr_td_eb . 'HM channel type' . $dthdr_td_eb . 'Dir.' . $dthdr_td_eb
                 . 'AES' . $dthdr_td_e;
        $html .= '<td style="width: 2%; text-align: center; color: #EEEEEE; font-size: medium">Roa- ming</td>';

        foreach ($data['hm_BidCos_Ifc_list'] as $hm_ifce) {
            if ($hm_ifce['CONNECTED']) {
                $html .= '<td style="width: 6%; text-align: center; color: #EEEEEE; font-size: small">' . $hm_ifce['ADDRESS']
                         . ' tx/rx&nbsp;(db&micro;V)</td>';
            }
        }
        $html .= '</tr>' . PHP_EOL;

        $entry_no        = 0;
        $previous_hm_adr = '';
        foreach ($data['HM_array'] as $HM_dev) {
            $hm_adr          = explode(':', $HM_dev['HM_address']);
            $same_device     = ($hm_adr[0] === $previous_hm_adr);
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
        $html .= '<td class="' . $td_color . ' text-right">' . $entryNo . '&nbsp;&nbsp;</td>';
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
        if (!isset($dev['HM_levels'])) {
            return str_repeat('<td></td>', $connectedIfcNum);
        }

        $html = '';
        for ($i = 0; $i < $connectedIfcNum; $i++) {
            if (isset($dev['HM_levels'][$i])) {
                $lciValue     = $dev['HM_levels'][$i];
                $isBest       = (bool)($lciValue[3] ?? false);
                $isAssociated = ($dev['HM_Roaming'] === '+') || ($lciValue[2] ?? false);

                if ($isBest) {
                    $color = $isAssociated ? '#DDDD66' : '#FFFF88';
                } else {
                    $color = '#DDDDDD';
                }

                [$rx, $tx] = $this->getRxTxLevelString((int)$lciValue[0], (int)$lciValue[1]);

                $content = $isAssociated ? "<ins>$rx &#047; $tx</ins>" : "$rx &#047; $tx";
                $html    .= '<td class="text-center"><p style="color: ' . $color . '">' . $content . '</p></td>';
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

    private function generateTableRow(string $fontSize, string $color, string $content): string
    {
        return "<tr><td style='font-size: $fontSize; color: $color'>$content</td>";
    }

    private function formatInterfaceRow(array $hm_ifce): string
    {
        $italicStart = $hm_ifce['DEFAULT'] ? '<i>' : '';
        $italicEnd   = $hm_ifce['DEFAULT'] ? '</i>' : '';

        $connectionStatus = $hm_ifce['CONNECTED'] ? 'connected' : 'Not connected';
        $interfaceInfo    = sprintf('%s (Fw: %s, DC: %s%%)', $hm_ifce['ADDRESS'], $hm_ifce['FIRMWARE_VERSION'], $hm_ifce['DUTY_CYCLE']);

        return "<tr>
                <td style='font-size: small;color: #EEEEEE'>{$italicStart}Interface: $interfaceInfo&nbsp;{$italicEnd}</td>
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
     *
     * @return string The generated HTML content.
     */
    private function getHtmlContent(
        string $HTML_intro,
        string $HTML_ifcs,
        string $HTML_sep,
        string $HTML_dvcs,
        string $HTML_notes
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

        .bgcolor-header-devices td {
            position: sticky;
            top: 0;
            background-color: #334455;
            z-index: 1;
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
  </table>
</body>
</html>
HEREDOC;
    }

    //
    // Versionsinformationen
    //

    /**
     * Retrieves the details of a library from a JSON file.
     *
     * @return array The details of the library as an associative array.
     * @throws \JsonException
     */
    private function getLibraryDetails(): array
    {
        $filename = __DIR__ . DIRECTORY_SEPARATOR . '..' . DIRECTORY_SEPARATOR . 'library.json';
        $content  = file_get_contents($filename);
        if ($content === false) {
            return [];
        }
        return json_decode($content, true, 512, JSON_THROW_ON_ERROR);
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
        return sprintf('%s.%s', $library['version'] ?? '?', $library['build'] ?? '?');
    }

    //
    // Gateway (CCU) und Geräte-Abfragen
    //

    private function fetchParentId(): int
    {
        return IPS_GetInstance($this->InstanceID)['ConnectionID'];
    }

    private function isGatewayActive(int $parentId): bool
    {
        if ($parentId === 0) {
            echo $this->Translate('Gateway is not configured!') . PHP_EOL . PHP_EOL;
            return false;
        }
        $this->SendDebug('Parent', sprintf('%s (#%s)', IPS_GetName($parentId), $parentId), 0);

        if (($this->GetStatus() !== IS_ACTIVE) || !$this->HasActiveParent()) {
            echo $this->Translate('Instance is not active!') . PHP_EOL . PHP_EOL;
            return false;
        }
        return true;
    }

    /**
     * Retrieves the device lists of all BidCos services (RF, IP, WR).
     *
     * @return array{0: array, 1: array<string, int>, 2: int} [Geräteliste, Gerätezähler je Service, Fehlerzähler]
     * @throws \JsonException
     */
    private function getDeviceLists(int $parentId, array $parentConfig): array
    {
        $ccuHost    = IPS_GetProperty($parentId, 'Host');
        $devLists   = [];
        $devCounter = [];
        $errorCount = 0;

        foreach (self::SERVICETYPES as $type) {
            $devCounter[$type] = 0;

            $response = $this->SendRequestMessage(
                'listDevices',
                [],
                $this->getBidCosServiceAddress($ccuHost, $parentConfig, $type),
                $parentConfig['UseSSL'],
                $parentConfig['Password'],
                $parentConfig['Username']
            );

            if ($response->errno !== 0) {
                $this->SendDebug('Error', sprintf(self::ERROR_MSG, $type), 0);
                $errorCount++;
                continue;
            }

            $devices = (new Encoder())->decode($response->value());
            foreach ($devices as $device) {
                if (($device['PARENT'] === '') && ($device['ADDRESS'] !== 'BidCoS-Wir')) {
                    $devCounter[$type]++;
                }
            }
            $devLists[] = $devices;
        }

        return [array_merge(...$devLists), $devCounter, $errorCount];
    }

    private function getBidCosServiceAddress(string $ccuHost, array $parentConfig, string $type): string
    {
        if ($parentConfig['UseSSL']) {
            return sprintf('https://%s:%s', $ccuHost, $parentConfig[$type . 'SSLPort']);
        }
        return sprintf('http://%s:%s', $ccuHost, $parentConfig[$type . 'Port']);
    }

    //
    // Registrierung, Status, Fortschrittsanzeige
    //

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
            $this->RegisterVariableString(
                self::VAR_IDENT_DEVICELIST,
                $this->Translate('Device List'),
                ['PRESENTATION' => VARIABLE_PRESENTATION_VALUE_PRESENTATION],
                1
            );
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

    private function progressBarInit(): void
    {
        $this->progressBarStep = 0;
        $this->UpdateFormField('ProgressBar', 'maximum', self::PROGRESS_BAR_MAX);
        $this->UpdateFormField('ProgressBar', 'visible', true);
    }

    private function advanceProgressBar(): void
    {
        $this->UpdateFormField('ProgressBar', 'current', $this->progressBarStep++);
    }

    //
    // Transport (XML-RPC und HM-Script)
    //

    /**
     * @throws \JsonException
     */
    private function SendRequestMessage(
        string $methodName,
        array $params,
        string $serviceAddress,
        bool $useSSL,
        string $password,
        string $username
    ): Response {
        $client = new Client($serviceAddress);
        if ($useSSL) {
            // Die CCU verwendet ein selbstsigniertes Zertifikat
            $client->setOption(Client::OPT_VERIFY_HOST, 0);
            $client->setOption(Client::OPT_VERIFY_PEER, false);
        }
        if ($password !== '') {
            $client->setCredentials($username, $password);
        }

        $this->SendDebug(
            'send (xmlrpc)',
            sprintf('send (xmlrpc):%s:%s, params: %s', $serviceAddress, $methodName, json_encode($params, JSON_THROW_ON_ERROR)),
            0
        );

        return $client->send(new Request($methodName, $params));
    }

    private function getHMChannelName(string $ccuHost, string $deviceAddress): string
    {
        $hmScript = 'Name = (xmlrpc.GetObjectByHSSAddress(interfaces.GetAt(0), "' . $deviceAddress . '")).Name();' . PHP_EOL;

        $scriptReturn = $this->SendScript($ccuHost, $hmScript);
        if ($scriptReturn === false) {
            return '';
        }

        $channelName = json_decode($scriptReturn, true, 512, JSON_THROW_ON_ERROR)['Name'] ?? '';
        if (!is_string($channelName)) { //Wenn der ChannelName auf HM Seite leer ist, kommt ein leeres Array zurück
            $channelName = '';
        }

        $this->SendDebug(__FUNCTION__, sprintf('HMAddress: %s, HMDeviceAddress: %s -> %s', $ccuHost, $deviceAddress, $channelName), 0);
        return $channelName;
    }

    private function SendScript(string $ccuHost, string $hmScript): false|string
    {
        try {
            $hmScriptResult = $this->LoadHMScript($ccuHost, 'Script.exe', $hmScript);
            $xml            = new SimpleXMLElement(
                mb_convert_encoding($hmScriptResult, 'UTF-8', 'ISO-8859-1'),
                LIBXML_NOBLANKS | LIBXML_NONET
            );
        } catch (Exception $exc) {
            $this->LogMessage($exc->getMessage(), KL_ERROR);
            return false;
        }

        unset($xml->exec, $xml->sessionId, $xml->httpUserAgent);
        return json_encode($xml, JSON_THROW_ON_ERROR);
    }

    /**
     * Loads an HM script from a given HM address using cURL.
     *
     * @param string $ccuHost  The CCU address.
     * @param string $url      The URL of the HM script to load.
     * @param string $hmScript The HM script to load.
     *
     * @return string The result of the cURL execution.
     * @throws \InvalidArgumentException
     * @throws \JsonException
     * @throws \RuntimeException If the request fails or the HTTP response code is >= 400.
     */
    private function LoadHMScript(string $ccuHost, string $url, string $hmScript): string
    {
        if ($ccuHost === '') {
            throw new InvalidArgumentException('CCU Address not set.');
        }

        $ch       = $this->setupCurlHandler($ccuHost, $url, $hmScript);
        $result   = curl_exec($ch);
        $httpCode = (int)curl_getinfo($ch, CURLINFO_HTTP_CODE);

        if (($result === false) || ($httpCode >= 400)) {
            throw new RuntimeException('CCU unreachable');
        }

        return $result;
    }

    /**
     * Sets up a cURL handler for making requests to the specified HM address.
     *
     * @param string $ccuHost  The CCU address.
     * @param string $url      The URL to send the request to.
     * @param string $hmScript The payload to send with the request.
     *
     * @return \CurlHandle The cURL handler.
     * @throws \JsonException
     * @throws \RuntimeException
     */
    private function setupCurlHandler(string $ccuHost, string $url, string $hmScript): CurlHandle
    {
        $parentId     = $this->fetchParentId();
        $parentConfig = json_decode(IPS_GetConfiguration($parentId), true, 512, JSON_THROW_ON_ERROR);
        $scheme       = $parentConfig['UseSSL'] ? 'https' : 'http';
        $port         = $parentConfig['UseSSL'] ? $parentConfig['HSSSLPort'] : $parentConfig['HSPort'];

        $ch = curl_init(sprintf('%s://%s:%s/%s', $scheme, $ccuHost, $port, $url));
        if ($ch === false) {
            throw new RuntimeException('curl_init failed');
        }

        $header = [
            'Accept: text/plain,text/xml,application/xml,application/xhtml+xml,text/html',
            'Cache-Control: max-age=0',
            'Connection: close',
            'Accept-Charset: UTF-8',
            'Content-type: text/plain;charset="UTF-8"',
            'Expect:'
        ];

        if ($parentConfig['UseSSL']) {
            curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
            curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, false);
        }

        curl_setopt($ch, CURLOPT_HEADER, false);
        curl_setopt($ch, CURLOPT_FAILONERROR, true);
        curl_setopt($ch, CURLOPT_POST, true);
        curl_setopt($ch, CURLOPT_POSTFIELDS, $hmScript);
        curl_setopt($ch, CURLOPT_HTTPHEADER, $header);
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
        curl_setopt($ch, CURLOPT_CONNECTTIMEOUT_MS, 1000);
        curl_setopt($ch, CURLOPT_HTTP_VERSION, CURL_HTTP_VERSION_1_1);
        curl_setopt($ch, CURLOPT_TIMEOUT_MS, 5000);

        if ($parentConfig['Password'] !== '') {
            curl_setopt($ch, CURLOPT_HTTPAUTH, CURLAUTH_BASIC);
            curl_setopt($ch, CURLOPT_USERPWD, $parentConfig['Username'] . ':' . $parentConfig['Password']);
        }

        return $ch;
    }

    //
    // Sortierung
    //

    private function sortDeviceRecords(array &$records): void
    {
        $sortField = match ($this->ReadPropertyInteger(self::PROP_SORTORDER)) {
            1       => 'HM_devtype',
            2       => 'HM_device',
            3       => 'IPS_name',
            4       => 'HM_devname',
            default => null // HM address
        };

        if ($sortField === null) {
            usort($records, self::compareByAddress(...));
        } else {
            usort($records, static fn(array $a, array $b): int => self::compareByField($a, $b, $sortField));
        }
    }

    private static function compareByField(array $a, array $b, string $field): int
    {
        $result = strcasecmp((string)$a[$field], (string)$b[$field]);
        if ($result === 0) {
            $result = self::compareByAddress($a, $b);
        }
        return $result;
    }

    /**
     * Vergleicht zwei Geräteeinträge anhand der HM-Adresse (Groß-/Kleinschreibung wird ignoriert).
     * Bei Adressen der Form "xxx:yyy" mit gleichem Geräteteil wird die Kanalnummer verglichen.
     */
    private static function compareByAddress(array $a, array $b): int
    {
        $result = strcasecmp($a['HM_address'], $b['HM_address']);

        $a_adr = explode(':', $a['HM_address']);
        $b_adr = explode(':', $b['HM_address']);
        if (count($a_adr) === 2 && count($b_adr) === 2 && strcasecmp($a_adr[0], $b_adr[0]) === 0) {
            $result = (int)($a_adr[1] > $b_adr[1]);
        }

        return $result;
    }
}
