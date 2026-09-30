<?php

declare(strict_types=1);

/*
 * Gemeinsamer Testrahmen: bindet HMInventoryReportCreator an den offiziellen Kernel-Stub
 * (symcon/SymconStubs, Submodul tests/stubs, gepinnt) und macht die privaten Methoden für Tests
 * aufrufbar. Netzzugriffe (XML-RPC, HM-Script) finden in den getesteten Methoden nicht statt.
 *
 * Einbinden mit require_once __DIR__ . '/harness.php'; Instanzen über neueInstanz().
 * Vorlage: T:/modules/SymconJuControlConnector/tests/harness.php
 */

require_once __DIR__ . '/stubs/autoload.php';

// Warnungen und Notices des Moduls sollen Tests abbrechen, nicht still durchlaufen.
// E_USER_NOTICE (Stub meldet so einen unbekannten Ident) und E_DEPRECATED bleiben außen vor.
set_error_handler(static function (int $nr, string $text, string $datei, int $zeile): bool {
    if (!(error_reporting() & $nr)) {
        return false;
    }
    if ($nr & (E_USER_ERROR | E_USER_WARNING | E_WARNING | E_NOTICE)) {
        throw new ErrorException($text, 0, $nr, $datei, $zeile);
    }
    return false;
});

require_once dirname(__DIR__) . '/HM_Inventory/module.php';

final class HMInventoryHarness extends HMInventoryReportCreator
{
    public const MODULE_ID = '{E3BEF9D8-23D4-47A8-B823-53BD7AF65CC3}'; // HM_Inventory/module.json

    /** Der Stub verlangt eine Uhr für RegisterTimer/SetTimerInterval. */
    protected function getTime(): int
    {
        return 1790780000; // fest: 30.09.2026
    }

    /** Eine Eigenschaft setzen und übernehmen, wie ein Speichern des Formulars. */
    public function einstellen(string $name, mixed $wert): void
    {
        $this->SetProperty($name, $wert);
        $this->ApplyChanges();
    }

    /** Private Methode des Moduls aufrufen; existiert sie nicht, wirft ReflectionException. */
    public function rufe(string $methode, mixed ...$argumente): mixed
    {
        $m = new ReflectionMethod(HMInventoryReportCreator::class, $methode);
        return $m->isStatic() ? $m->invokeArgs(null, $argumente) : $m->invokeArgs($this, $argumente);
    }

    /** Private Methode mit Referenzparameter (z. B. sortDeviceRecords(array &$records)). */
    public function rufeMitReferenz(string $methode, array &$argument): void
    {
        $m = new ReflectionMethod(HMInventoryReportCreator::class, $methode);
        $m->invokeArgs($this, [&$argument]);
    }
}

/** Legt eine Instanz im Kernel-Stub an (Create + ApplyChanges laufen in createInstance). */
function neueInstanz(): HMInventoryHarness
{
    $id = IPS\ObjectManager::registerObject(1 /* Instance */);
    IPS\InstanceManager::createInstance($id, [
        'ModuleID'   => HMInventoryHarness::MODULE_ID,
        'ModuleName' => 'HM Inventory Report Creator',
        'ModuleType' => 3,
        'Class'      => HMInventoryHarness::class,
    ]);
    return IPS\InstanceManager::getInstanceInterface($id);
}

/** Echte Geräteliste (Variable DeviceList der nuc-Instanz, Namen geleert). */
function fixtureGeraeteliste(): array
{
    return json_decode(file_get_contents(__DIR__ . '/fixtures/devicelist_nuc_2026-09-30.json'), true, 512, JSON_THROW_ON_ERROR);
}

$pruefungen = 0;
$fehler     = [];
function pruefe(bool $ok, string $text): void
{
    global $pruefungen, $fehler;
    $pruefungen++;
    if (!$ok) {
        $fehler[] = $text;
    }
    echo ($ok ? '  ok   ' : '  FEHL ') . $text . "\n";
}

/** Schlusszeile und Exit-Code */
function ergebnis(): never
{
    global $pruefungen, $fehler;
    echo "\n$pruefungen Prüfungen, " . count($fehler) . " Fehler\n";
    exit($fehler === [] ? 0 : 1);
}

IPS\Kernel::reset(); // einmal je Testlauf
