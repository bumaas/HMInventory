<?php

declare(strict_types=1);

/*
 * Der HTML-Report ist in der Systemsprache von Symcon beschriftet (Hinweis aus dem Store-Review der
 * Stable 2.2 #28, 10/2026: „Der in HTML generierte Bericht wird scheinbar nicht übersetzt.“).
 *
 * Die Harness übersetzt Translate() wie ein deutsches Symcon (locale.json, translations.de). Jede
 * Report-Sektion wird mit der echten Geräteliste des nuc (tests/fixtures/devicelist_nuc_2026-09-30.json)
 * gerendert; danach müssen die deutschen Beschriftungen drinstehen und keines der bisher fest
 * verdrahteten englischen Literale mehr auftauchen. Werte aus den Daten (TX/RX, +/-, dBm, Adressen,
 * Typen) sind keine Beschriftung und bleiben, wie sie sind.
 *
 * Aufruf: php tests/check-report-uebersetzung.php
 */

require_once __DIR__ . '/harness.php';

$m     = neueInstanz();
$liste = fixtureGeraeteliste();

// Interfaces aus der Fixture: das erste als Standard-Interface, eines davon nicht verbunden
$interfaces = [];
foreach ($liste as $e) {
    $interfaces[$e['HM_Interface']] = ['ADDRESS' => $e['HM_Interface'], 'DESCRIPTION' => 'Fixture', 'FIRMWARE_VERSION' => '1.0.0', 'DUTY_CYCLE' => 7, 'CONNECTED' => true, 'DEFAULT' => false];
}
$interfaces = array_values($interfaces);
$interfaces[0]['DEFAULT'] = true;
$interfaces[] = ['ADDRESS' => 'OFFLINE0001', 'DESCRIPTION' => 'getrennt', 'FIRMWARE_VERSION' => '0.0.0', 'DUTY_CYCLE' => 0, 'CONNECTED' => false, 'DEFAULT' => false];
$verbunden  = count($interfaces) - 1;

$daten = [
    'dev_counter'                => ['RF' => 10, 'WR' => 1, 'IP' => 20],
    'HM_array'                   => $liste,
    'hm_BidCos_Ifc_list'         => $interfaces,
    'HM_interface_num'           => count($interfaces),
    'HM_interface_connected_num' => $verbunden,
    'IPS_device_num'             => 5,
    'IPS_HM_channel_num'         => 6,
    'HM_module_num'              => count($liste),
    'HM_default_interface_no'    => 0,
];

$m->einstellen('ShowHMConfiguratorDeviceNames', true);
$kopf       = $m->rufe('renderHeaderSection', $daten);
$ifcs       = $m->rufe('renderInterfacesSection', $interfaces);
$geraete    = $m->rufe('renderDevicesSection', $daten);
$leer       = $m->rufe('renderDevicesSection', array_merge($daten, ['HM_array' => [], 'HM_module_num' => 0]));
$hinweise   = $m->rufe('renderNotesSection');
$dokument   = $m->rufe('getHtmlContent', $kopf, $ifcs, $m->rufe('renderSeparator'), $geraete, $hinweise);
$alles      = $kopf . $ifcs . $geraete . $leer . $hinweise;

echo "Deutsche Beschriftung (Systemsprache de_DE)\n";
$erwartet = [
    'Kopf: Erstellungszeitpunkt'      => [$kopf, 'erstellt am'],
    'Kopf: Interface-Zähler'          => [$kopf, 'HomeMatic-Interfaces'],
    'Kopf: Instanz-Zähler'            => [$kopf, 'Instanzen'],
    'Interfaces: verbunden'           => [$ifcs, '>verbunden<'],
    'Interfaces: nicht verbunden'     => [$ifcs, 'nicht verbunden'],
    'Spalte „HM Adresse“'             => [$geraete, 'HM Adresse'],
    'Spalte „HM Gerätetyp“'           => [$geraete, 'HM Gerätetyp'],
    'Spalte „HM Kanaltyp“'            => [$geraete, 'HM Kanaltyp'],
    'Spalte „HM Gerätename“'          => [$geraete, 'HM Gerätename'],
    'Spalte „IPS Gerätename“'         => [$geraete, 'IPS Gerätename'],
    'Spalte „Roaming“ ohne Trennstrich' => [$geraete, 'Roaming'],
    'Leerer Report: Hinweis deutsch'  => [$leer, 'Keine HomeMatic-Geräte gefunden'],
    'Legende: Überschrift'            => [$hinweise, 'Hinweise:'],
    'Legende: Standard-Interface kursiv' => [$hinweise, 'kursiv'],
    'Legende: Pegelpaare'             => [$hinweise, 'Pegelpaar'],
    'Legende: rote Schrift'           => [$hinweise, 'rot'],
    'Dokument: lang="de"'             => [$dokument, '<html lang="de">'],
    'Dokument: Titel gesetzt'         => [$dokument, '<title>HM Inventory</title>'],
];
foreach ($erwartet as $text => [$html, $muster]) {
    pruefe(str_contains($html, $muster), sprintf('%s („%s“)', $text, $muster));
}

echo "\nKeine englischen Literale mehr\n";
$englisch = [
    'found at', 'HomeMatic interfaces', 'IPS instances', 'connected to',
    'Not connected', '>connected<',
    'IPS device name', 'HM address', 'HM device name', 'HM device type', 'HM channel type',
    'Roa- ming', 'No HomeMatic devices found',
    'Notes:', 'italic letters', 'Level-pairs', 'shown in red', 'haven\'t sent',
];
foreach ($englisch as $muster) {
    pruefe(!str_contains($alles, $muster), sprintf('„%s“ kommt nicht mehr vor', $muster));
}

echo "\nUnverändert\n";
pruefe(str_contains($geraete, '(dBm)'), 'Pegelspalten weiterhin in dBm');
pruefe(str_contains($geraete, '>TX<') && str_contains($geraete, '>RX<'), 'Richtungswerte TX/RX bleiben');
pruefe(substr_count($geraete, '<tr class="bg_color_') === count($liste), sprintf('je Fixture-Eintrag eine Zeile (%d)', count($liste)));
pruefe(str_contains($dokument, '<h1>HM Inventory</h1>'), 'Überschrift „HM Inventory“ bleibt (Produktname)');

ergebnis();
