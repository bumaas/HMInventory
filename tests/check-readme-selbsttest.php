<?php

declare(strict_types=1);

/**
 * Selbsttest für tests/check-readme.php.
 *
 * Fixtures sind die echten Dateien dieses Repos (README.md, HM_Inventory/*.json, module.php),
 * je Fall gezielt verändert und in ein Temp-Verzeichnis gelegt; dort läuft die Prüfung.
 * Die Lückenliste kommt über --bekannt-schreiben (ohne vorhandene readme-bekannt.json).
 *
 * Aufruf: php tests/check-readme-selbsttest.php [<pfad zu check-readme.php>]
 *         Vorgabe ist die Prüfung neben diesem Test; ein anderer Pfad dient dem Rot/Grün-Nachweis.
 */

$wurzel  = dirname(__DIR__);
$pruefer = $argv[1] ?? __DIR__ . '/check-readme.php';
$modul   = 'HM Inventory Report Creator';

$pruefungen = 0;
$fehler     = 0;
function pruefe(bool $ok, string $text): void
{
    global $pruefungen, $fehler;
    $pruefungen++;
    if (!$ok) {
        $fehler++;
    }
    echo ($ok ? '  ok   ' : '  FEHL ') . $text . "\n";
}

/**
 * Legt ein Fixture-Repo an: echte Dateien, README und Formular über Rückrufe verändert.
 *
 * @return string Verzeichnis
 */
function fixture(?callable $readme = null, ?callable $form = null, ?array $bekannt = null): string
{
    global $wurzel, $pruefer;
    $ziel = sys_get_temp_dir() . '/hmi-readme-' . bin2hex(random_bytes(4));
    mkdir("$ziel/HM_Inventory", 0777, true);
    mkdir("$ziel/tests");
    foreach (['module.json', 'form.json', 'locale.json', 'module.php'] as $datei) {
        copy("$wurzel/HM_Inventory/$datei", "$ziel/HM_Inventory/$datei");
    }
    register_shutdown_function(static fn () => exec((PHP_OS_FAMILY === 'Windows' ? 'rmdir /s /q ' : 'rm -rf ') . escapeshellarg($ziel)));
    copy($pruefer, "$ziel/tests/check-readme.php");
    $text = file_get_contents("$wurzel/README.md");
    file_put_contents("$ziel/README.md", $readme ? $readme($text) : $text);
    if ($form) {
        $f = json_decode(file_get_contents("$ziel/HM_Inventory/form.json"), true, 512, JSON_THROW_ON_ERROR);
        file_put_contents("$ziel/HM_Inventory/form.json", json_encode($form($f), JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES));
    }
    if ($bekannt !== null) {
        file_put_contents("$ziel/tests/readme-bekannt.json", json_encode(['offen' => $bekannt]));
    }
    return $ziel;
}

/** @return array{0: int, 1: string} Exit-Code und Ausgabe */
function lauf(string $ziel, string $schalter = ''): array
{
    exec(escapeshellarg(PHP_BINARY) . ' ' . escapeshellarg("$ziel/tests/check-readme.php") . " $schalter 2>&1", $aus, $code);
    return [$code, implode("\n", $aus)];
}

/** Lückenliste der Prüfung (--bekannt-schreiben in ein Fixture ohne readme-bekannt.json). */
function luecken(string $ziel): array
{
    lauf($ziel, '--bekannt-schreiben');
    $datei = "$ziel/tests/readme-bekannt.json";
    return is_file($datei) ? (json_decode(file_get_contents($datei), true)['offen'] ?? []) : ['(keine Datei geschrieben)'];
}

/** Zeilen der IP-Symcon-Prüfung: true = alle bestanden. */
function ipSymconOk(string $ziel): bool
{
    [, $aus] = lauf($ziel);
    $gefunden = false;
    foreach (explode("\n", $aus) as $zeile) {
        if (stripos($zeile, 'IP-Symcon') !== false) {
            if (!str_starts_with(ltrim($zeile), 'ok')) {
                return false;
            }
            $gefunden = true;
        }
    }
    return $gefunden;
}

/** Ersetzt genau einen Treffer - findet das Muster nichts, ist die Fixture veraltet (lauter Abbruch statt stillem Grün). */
function einmal(string $muster, string $ersatz, string $t): string
{
    $neu = preg_replace($muster, $ersatz, $t, 1, $anzahl);
    if ($anzahl !== 1) {
        throw new RuntimeException("Fixture: Muster $muster nicht im README gefunden");
    }
    return $neu;
}

/** Entfernt den Abschnitt einer Funktion aus der Funktionsreferenz und jede weitere Erwähnung (Beispiele). */
function ohneFunktion(string $funktion): callable
{
    return static fn (string $t): string => str_replace($funktion . '(', 'Beispiel(', einmal('/```php\R' . $funktion . '\(.*?```\R.*?\R/s', '', $t));
}

/** Entfernt die Tabellenzeile einer Eigenschaft (erste Spalte „`Name` (Beschriftung)“). */
function ohneZeile(string $feld): callable
{
    return static fn (string $t): string => einmal('/^\|\s*`?' . $feld . '`?[^|]*\|.*\R/m', '', $t);
}

echo 'Selbsttest check-readme.php (' . basename(dirname($pruefer)) . '/' . basename($pruefer) . ")\n";

// 1. Das echte README hat keine Lücken: Alle Felder stehen in der Tabelle (ohne Backticks).
$l = luecken(fixture());
pruefe($l === [], 'echtes README: keine Lücken' . ($l ? ' - gemeldet: ' . implode(', ', $l) : ''));

// 2./3. Funktionen, die ein Formularknopf aufruft, sind trotzdem Skript-API.
foreach (['HMI_CreateReport', 'HMI_GetReportUrl', 'HMI_GetOutputFileAbsolutePath'] as $fn) {
    $l = luecken(fixture(ohneFunktion($fn)));
    pruefe(in_array("funktion:$fn", $l, true), "$fn aus der Funktionsreferenz entfernt: als Lücke gemeldet");
}

// 4. Die Übersetzung „Ausgabedatei“ im Fließtext ersetzt keine Beschreibung des Feldes.
$l = luecken(fixture(ohneZeile('OutputFile')));
pruefe(in_array("feld:$modul:OutputFile", $l, true), 'Zeile OutputFile entfernt: als Lücke gemeldet');

// 5. „active“ zählt nicht als Teil eines anderen Wortes.
$l = luecken(fixture(static fn (string $t): string => einmal(
    '/^## 2\. /m',
    "Die Ausgabe ist interactive.\n\n$0",
    // übrige Nennungen entschärfen: `active` im Fließtext, „aktiv“ in den Meldungstabellen
    str_replace(['`active`', '| aktiv |', 'nicht aktiv!'], ['die Instanz', '| in Betrieb |', 'nicht bereit!'], ohneZeile('active')($t))
)));
pruefe(in_array("feld:$modul:active", $l, true), 'Zeile active entfernt, „interactive“ im Text: als Lücke gemeldet');

// 6. Ein einzelnes „<“ verschluckt keinen Text bis zum nächsten „>“.
// Anker vor der Feldtabelle; dahinter steht in der Zeile OutputFile ein „>“ (`<Symcon-Verzeichnis>`).
$lt = static fn (string $t): string => einmal(
    '/^\*\*In Symcon\*\*\R/m',
    "$0\nSchwacher Empfang: RSSI < -90 dBm. Gilt auch unter IP-Symcon.\n",
    $t
);
$z = fixture($lt);
$l = luecken($z);
pruefe($l === [], '„RSSI < -90“ vor der Tabelle: Tabelle weiter gelesen' . ($l ? ' - gemeldet: ' . implode(', ', $l) : ''));
pruefe(!ipSymconOk(fixture($lt)), '„IP-Symcon“ hinter einem einzelnen „<“ wird gefunden');

// 7. Linktext zählt nicht, Groß-/Kleinschreibung schon.
pruefe(ipSymconOk(fixture(static fn (string $t): string => $t . "\nSiehe [IP-Symcon Forum](https://community.symcon.de/t/1).\n")), 'Linktext [IP-Symcon …](…) wird nicht beanstandet');
pruefe(!ipSymconOk(fixture(static fn (string $t): string => $t . "\n## HINWEIS FÜR IP-SYMCON\n")), '„IP-SYMCON“ in Großbuchstaben wird gefunden');

// 8. Auch sichtbare Formulartexte werden geprüft.
pruefe(!ipSymconOk(fixture(null, static function (array $f): array {
    $f['elements'][] = ['type' => 'Label', 'caption' => 'Only for IP-Symcon'];
    return $f;
})), '„IP-Symcon“ in einer form.json-Beschriftung wird gefunden');

// 9. --bekannt-schreiben ergänzt keine neuen Lücken, streicht aber geschlossene.
$z                  = fixture(ohneZeile('SortOrder'), null, []);
[$code]             = lauf($z, '--bekannt-schreiben');
$nachher            = json_decode(file_get_contents("$z/tests/readme-bekannt.json"), true)['offen'] ?? null;
pruefe($code !== 0 && $nachher === [], '--bekannt-schreiben bei neuer Lücke: abgelehnt, Liste unverändert');
$z       = fixture(ohneZeile('SortOrder'), null, ["feld:$modul:SortOrder", "feld:$modul:UpdateInterval"]);
[$code]  = lauf($z, '--bekannt-schreiben');
$nachher = json_decode(file_get_contents("$z/tests/readme-bekannt.json"), true)['offen'] ?? null;
pruefe($code === 0 && $nachher === ["feld:$modul:SortOrder"], '--bekannt-schreiben streicht geschlossene Lücken');

echo "\n$pruefungen Prüfungen, $fehler Fehler\n";
exit($fehler === 0 ? 0 : 1);
