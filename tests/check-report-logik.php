<?php

declare(strict_types=1);

/*
 * Report-Logik gegen die echte Geräteliste des nuc (tests/fixtures/devicelist_nuc_2026-09-30.json):
 *
 * 1. Sortierung: Jede Option aus form.json sortiert nach dem Feld, das im Report unter der
 *    gleichnamigen Spalte steht (HM device type = HM_device, HM channel type = HM_devtype).
 * 2. Gelbe Markierung (HM_levels[i][3]): immer der höchste rechte Wert (wie gut das Interface das
 *    Gerät hört) - ein Maßstab für alle Geräte, auch wenn links Werte bekannt sind; ohne rechten
 *    Wert keine Markierung.
 *    65536 steht für „kein Wert“.
 * 3. Beschriftung: Pegel in dBm (nicht dbµV), Standard-Interface in der Legende „kursiv“ (so
 *    wird es dargestellt), nicht „fett“.
 *
 * Aufruf: php tests/check-report-logik.php
 */

require_once __DIR__ . '/harness.php';

const KEIN_WERT = 65536;

$m     = neueInstanz();
$liste = fixtureGeraeteliste();

echo "Sortierung\n";
// Option (Wert in form.json) => [Beschriftung, Feld, das der Report unter dieser Spalte zeigt]
$optionen = [
    1 => ['HM device type', 'HM_device'],
    2 => ['HM channel type', 'HM_devtype'],
    3 => ['IPS device name', 'IPS_name'],
    4 => ['HM device name', 'HM_devname'],
];
foreach ($optionen as $wert => [$beschriftung, $feld]) {
    $m->einstellen('SortOrder', $wert);
    $records = array_reverse($liste);
    $m->rufeMitReferenz('sortDeviceRecords', $records);
    $falsch = 0;
    for ($i = 1, $n = count($records); $i < $n; $i++) {
        if (strcasecmp((string)$records[$i - 1][$feld], (string)$records[$i][$feld]) > 0) {
            $falsch++;
        }
    }
    pruefe($falsch === 0, sprintf('„%s“ sortiert nach %s (%d Fehlstellungen)', $beschriftung, $feld, $falsch));
}
$m->einstellen('SortOrder', 0);
$records = array_reverse($liste);
$m->rufeMitReferenz('sortDeviceRecords', $records);
$falsch = 0;
for ($i = 1, $n = count($records); $i < $n; $i++) {
    if (strcasecmp(explode(':', $records[$i - 1]['HM_address'])[0], explode(':', $records[$i]['HM_address'])[0]) > 0) {
        $falsch++;
    }
}
pruefe($falsch === 0, sprintf('„HM address“ sortiert nach Geräteadresse (%d Fehlstellungen)', $falsch));

$mitPegeln = array_values(array_filter($liste, static fn (array $e): bool => isset($e['HM_levels'])));
printf("\nGelbe Markierung (%d Einträge mit Pegeln)\n", count($mitPegeln));
$gueltig = static fn (array $werte): array => array_filter($werte, static fn (int $v): bool => $v !== KEIN_WERT);
$zaehler = ['mit linkem Wert' => 0, 'nur rechter Wert' => 0, 'links anders als rechts' => 0];
$falsch  = ['mit linkem Wert' => [], 'nur rechter Wert' => [], 'unverändert' => []];
try {
    foreach ($mitPegeln as $e) {
        $eingabe  = array_map(static fn (array $p): array => [$p[0], $p[1], $p[2], false], $e['HM_levels']);
        $ergebnis = $m->rufe('markBestInterface', $eingabe);

        $markiert = array_keys(array_filter($ergebnis, static fn (array $p): bool => $p[3] === true));
        $links    = $gueltig(array_column($eingabe, 0));
        $rechts   = $gueltig(array_column($eingabe, 1));
        $erwartet = $rechts === [] ? [] : [array_search(max($rechts), $rechts, true)];
        $fall     = $links !== [] ? 'mit linkem Wert' : 'nur rechter Wert';
        $zaehler[$fall]++;
        if ($links !== [] && $rechts !== [] && array_search(max($links), $links, true) !== $erwartet[0]) {
            $zaehler['links anders als rechts']++;
        }
        if ($markiert !== $erwartet) {
            $falsch[$fall][] = $e['HM_address'];
        }
        $ohneMarke = array_map(static fn (array $p): array => array_slice($p, 0, 3), $ergebnis);
        if ($ohneMarke !== array_map(static fn (array $p): array => array_slice($p, 0, 3), $eingabe)) {
            $falsch['unverändert'][] = $e['HM_address'];
        }
    }
    $beispiel = static fn (array $a): string => $a === [] ? '' : ' - z. B. ' . implode(', ', array_slice($a, 0, 3));
    foreach (['mit linkem Wert', 'nur rechter Wert'] as $fall) {
        pruefe($falsch[$fall] === [], sprintf('%s: bester rechter Wert markiert (%d von %d falsch%s)', $fall, count($falsch[$fall]), $zaehler[$fall], $beispiel($falsch[$fall])));
    }
    pruefe($falsch['unverändert'] === [], 'Pegel und Zuordnung bleiben unverändert');
    pruefe($zaehler['links anders als rechts'] > 0, sprintf('Fixture enthält Einträge, bei denen links ein anderes Interface vorn liegt (%d)', $zaehler['links anders als rechts']));

    // Ohne rechten Wert keine Markierung - auch nicht nach dem linken. Aus einem echten Eintrag mit
    // linken Werten abgeleitet, rechts geleert (die Anlage liefert rechts immer einen Wert).
    $vorlage = null;
    foreach ($mitPegeln as $e) {
        if ($gueltig(array_column($e['HM_levels'], 0)) !== []) {
            $vorlage = $e;
            break;
        }
    }
    $ohneRechts = array_map(static fn (array $p): array => [$p[0], KEIN_WERT, $p[2], false], $vorlage['HM_levels']);
    $markiert   = array_filter($m->rufe('markBestInterface', $ohneRechts), static fn (array $p): bool => $p[3] === true);
    pruefe($markiert === [], "ohne rechten Wert: nichts markiert ({$vorlage['HM_address']}, rechts geleert)");
} catch (ReflectionException $e) {
    pruefe(false, 'Markierung als eigene Methode markBestInterface prüfbar (' . $e->getMessage() . ')');
}

echo "\nBeschriftung\n";
$interfaces = [];
foreach ($liste as $e) {
    $interfaces[$e['HM_Interface']] = ['ADDRESS' => $e['HM_Interface'], 'CONNECTED' => true, 'DEFAULT' => false];
}
$html = $m->rufe('renderDevicesSection', [
    'hm_BidCos_Ifc_list'         => array_values($interfaces),
    'HM_array'                   => [],
    'HM_interface_connected_num' => count($interfaces),
    'HM_module_num'              => 0,
]);
pruefe(str_contains($html, 'dBm'), 'Pegelspalten in dBm beschriftet');
pruefe(!str_contains($html, '&micro;V') && !str_contains($html, 'µV'), 'keine Angabe „dbµV“ mehr');
$notes = $m->rufe('renderNotesSection');
pruefe(str_contains($notes, 'italic') && !str_contains($notes, 'bold letters'), 'Legende: Standard-Interface kursiv, wie dargestellt');
pruefe(!str_contains($m->rufe('formatInterfaceRow', ['ADDRESS' => 'X', 'FIRMWARE_VERSION' => '1', 'DUTY_CYCLE' => 0, 'DESCRIPTION' => '', 'CONNECTED' => true, 'DEFAULT' => true]), '<b>'), 'Standard-Interface wird kursiv dargestellt (nicht fett)');

ergebnis();
