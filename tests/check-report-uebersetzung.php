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
 * Dazu die Befunde des Code-Reviews zu build 34: Zeichensatz im Dokument, Sprachcode aus der
 * Übersetzung statt aus der Systemsprache, Zeitstempel im Format der Sprache, „Symcon“ statt „IPS“
 * in neuem deutschem Text, Roaming-Spalte mit weichem Trennstrich statt overflow-wrap.
 *
 * Aufruf: php tests/check-report-uebersetzung.php
 */

require_once __DIR__ . '/harness.php';

$m     = neueInstanz();
$liste = fixtureGeraeteliste();

// Interfaces aus der Fixture: das erste als Standard-Interface, dazu eines, das nicht verbunden ist
$interfaces = fixtureInterfaces($liste);
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

/** Rendert alle Sektionen in der eingestellten Harness-Sprache. */
function rendern(HMInventoryHarness $m, array $daten, array $interfaces): array
{
    $kopf     = $m->rufe('renderHeaderSection', $daten);
    $ifcs     = $m->rufe('renderInterfacesSection', $interfaces);
    $geraete  = $m->rufe('renderDevicesSection', $daten);
    $leer     = $m->rufe('renderDevicesSection', array_merge($daten, ['HM_array' => [], 'HM_module_num' => 0]));
    $hinweise = $m->rufe('renderNotesSection');
    $dokument = $m->rufe('getHtmlContent', $kopf, $ifcs, $m->rufe('renderSeparator'), $geraete, $hinweise);
    return [$kopf, $ifcs, $geraete, $leer, $hinweise, $dokument];
}

$m->einstellen('ShowHMConfiguratorDeviceNames', true);
HMInventoryHarness::$sprache = 'de';
[$kopf, $ifcs, $geraete, $leer, $hinweise, $dokument] = rendern($m, $daten, $interfaces);
$alles = $kopf . $ifcs . $geraete . $leer . $hinweise;

echo "Deutsche Beschriftung (Übersetzung de)\n";
$erwartet = [
    'Kopf: Erstellungszeitpunkt'      => [$kopf, 'erstellt am'],
    'Kopf: Interface-Zähler'          => [$kopf, 'HomeMatic-Interfaces'],
    'Kopf: Instanz-Zähler'            => [$kopf, 'Symcon-Instanzen'],
    'Interfaces: verbunden'           => [$ifcs, '>verbunden<'],
    'Interfaces: nicht verbunden'     => [$ifcs, 'nicht verbunden'],
    'Spalte „HM Adresse“'             => [$geraete, 'HM Adresse'],
    'Spalte „HM Gerätetyp“'           => [$geraete, 'HM Gerätetyp'],
    'Spalte „HM Kanaltyp“'            => [$geraete, 'HM Kanaltyp'],
    'Spalte „HM Gerätename“'          => [$geraete, 'HM Gerätename'],
    'Spalte „IPS Gerätename“'         => [$geraete, 'IPS Gerätename'],
    'Spalte „Roaming“ mit weichem Trennstrich' => [$geraete, 'Roa&shy;ming'],
    'Leerer Report: Hinweis deutsch'  => [$leer, 'Keine HomeMatic-Geräte gefunden'],
    'Legende: Überschrift'            => [$hinweise, 'Hinweise:'],
    'Legende: Standard-Interface kursiv' => [$hinweise, 'kursiv'],
    'Legende: Pegelpaare'             => [$hinweise, 'Pegelpaar'],
    'Legende: rote Schrift'           => [$hinweise, 'rot'],
    'Legende: Symcon-Instanz'         => [$hinweise, 'Symcon-Instanz'],
    'Dokument: lang="de"'             => [$dokument, '<html lang="de">'],
    'Dokument: Titel gesetzt'         => [$dokument, '<title>HM Inventory</title>'],
    'Dokument: DOCTYPE'               => [$dokument, '<!DOCTYPE html>'],
    'Dokument: Zeichensatz UTF-8'     => [$dokument, '<meta charset="utf-8">'],
];
foreach ($erwartet as $text => [$html, $muster]) {
    pruefe(str_contains($html, $muster), sprintf('%s („%s“)', $text, $muster));
}
pruefe(
    preg_match('/erstellt am \d{2}\.\d{2}\.\d{4} \d{2}:\d{2}:\d{2}/', $kopf) === 1,
    'Kopf: Zeitstempel deutsch (dd.mm.yyyy hh:mm:ss)'
);

echo "\nKeine englischen Literale mehr\n";
$englisch = [
    'found at', 'created at', 'HomeMatic interfaces', 'IPS instances', 'connected to',
    'Not connected', '>connected<',
    'IPS device name', 'HM address', 'HM device name', 'HM device type', 'HM channel type',
    'Roa- ming', 'No HomeMatic devices found',
    'Notes:', 'italic letters', 'Level-pairs', 'shown in red', 'haven\'t sent',
];
foreach ($englisch as $muster) {
    pruefe(!str_contains($alles, $muster), sprintf('„%s“ kommt nicht mehr vor', $muster));
}

echo "\nNeuer deutscher Text sagt „Symcon“, nicht „IPS“\n";
pruefe(!str_contains($kopf, 'IPS-Instanzen'), 'Kopf: kein „IPS-Instanzen“');
pruefe(!str_contains($hinweise, 'IPS-Instanz'), 'Legende: kein „IPS-Instanz“');

echo "\nRoaming-Spalte\n";
pruefe(!str_contains($geraete, 'overflow-wrap'), 'kein overflow-wrap (zerlegt die 2%-Spalte in Einzelbuchstaben)');

echo "\nPegelspalten je verbundenem Interface\n";
pruefe(substr_count($geraete, '&nbsp;(dBm)') === $verbunden, sprintf('eine dBm-Kopfzelle je verbundenem Interface (%d)', $verbunden));
pruefe(!str_contains($geraete, '">&nbsp;(dBm)'), 'keine dBm-Kopfzelle mit leerer Adresse (Phantom-Interface aus HmIP-Einträgen)');

echo "\nUnverändert\n";
pruefe(str_contains($geraete, '(dBm)'), 'Pegelspalten weiterhin in dBm');
pruefe(str_contains($geraete, '>TX<') && str_contains($geraete, '>RX<'), 'Richtungswerte TX/RX bleiben');
pruefe(substr_count($geraete, '<tr class="bg_color_') === count($liste), sprintf('je Fixture-Eintrag eine Zeile (%d)', count($liste)));
pruefe(str_contains($dokument, '<h1>HM Inventory</h1>'), 'Überschrift „HM Inventory“ bleibt (Produktname)');

echo "\nEnglisch (keine Übersetzung für en)\n";
HMInventoryHarness::$sprache = 'en';
[$kopfEn, , , , $hinweiseEn, $dokumentEn] = rendern($m, $daten, $interfaces);
pruefe(str_contains($dokumentEn, '<html lang="en">'), 'Dokument: lang="en"');
pruefe(
    preg_match('/created at \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}/', $kopfEn) === 1,
    'Kopf: „created at“ mit Zeitstempel yyyy-mm-dd hh:mm:ss'
);
pruefe(!str_contains($kopfEn, 'found at'), 'Kopf: „found at“ auch im englischen Schlüssel ersetzt');
pruefe(str_contains($hinweiseEn, 'Notes:'), 'Legende: englisch');

echo "\nSprache ohne Übersetzung (fr): Dokument bleibt englisch\n";
HMInventoryHarness::$sprache = 'fr';
[, , , , , $dokumentFr] = rendern($m, $daten, $interfaces);
pruefe(str_contains($dokumentFr, '<html lang="en">'), 'Dokument: lang="en", nicht die Systemsprache');
HMInventoryHarness::$sprache = 'de';

ergebnis();