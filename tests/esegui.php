<?php

declare(strict_types=1);

// Prove del pacchetto, senza Composer e senza PHPUnit: `php tests/esegui.php`.
// Servono le estensioni zip e mbstring: senza zip il pacchetto ritorna null e le prove lo dicono.
require __DIR__ . '/../autoload.php';

use Dodaroe\FogliLeggeri\Excel;
use Dodaroe\FogliLeggeri\Word;
use Dodaroe\FogliLeggeri\Xml;

$ok = 0;
$ko = 0;
function prova(bool $condizione, string $nome): void
{
    global $ok, $ko;
    if ($condizione) {
        $ok++;
        echo "  ok  $nome\n";
    } else {
        $ko++;
        echo "  KO  $nome\n";
    }
}
function sezione(string $t): void
{
    echo "\n== $t ==\n";
}

/** Il contenuto di un file dentro lo zip (.xlsx e .docx sono zip). */
function dentro(string $file, string $percorso): string
{
    $zip = new ZipArchive();
    if ($zip->open($file) !== true) {
        return '';
    }
    $t = (string) $zip->getFromName($percorso);
    $zip->close();

    return $t;
}

if (!class_exists(ZipArchive::class)) {
    echo "Manca l'estensione zip di PHP: il pacchetto non puo creare file.\n";
    echo "Installala (su Debian/Ubuntu: php-zip) e rilancia.\n";
    exit(1);
}

sezione('Testo dentro gli XML di Office');

prova(Xml::testo("a & b <c>\x01") === 'a &amp; b &lt;c&gt;', 'escape di &, < e > e caratteri di controllo buttati');
prova(Xml::testo('"virgolette" e \'apici\'') === '&quot;virgolette&quot; e &apos;apici&apos;', 'virgolette e apici protetti (entita XML, non HTML)');
prova(Xml::testo(null) === '' && Xml::testo(12) === '12', 'null e numeri diventano testo');
prova(Xml::testo("riga\ncon\tandata a capo") === "riga\ncon\tandata a capo", 'a capo e tabulazione restano: sono ammessi');

sezione('Fogli di calcolo .xlsx');

$file = Excel::crea(['Elenco' => ['intestazioni' => ['Nome', 'CFU'], 'righe' => [['Ecologia', 6], ['Zoologia', 9]]]]);
prova($file !== null && is_file($file), 'il file si crea');
$foglio = dentro((string) $file, 'xl/worksheets/sheet1.xml');
prova(str_contains($foglio, '<autoFilter ref="A1:B3"/>'), 'intestazione con filtro sulle colonne giuste');
prova(str_contains($foglio, 'state="frozen"'), 'prima riga bloccata');
prova(str_contains($foglio, 'Ecologia') && str_contains($foglio, 'Zoologia'), 'i dati ci sono');
prova(str_contains(dentro((string) $file, 'xl/workbook.xml'), '<sheet name="Elenco"'), 'il foglio ha il nome dato');
prova(str_contains(dentro((string) $file, '[Content_Types].xml'), 'spreadsheetml.sheet.main+xml'), 'il pacchetto OOXML si dichiara come cartella di lavoro');
unlink((string) $file);

$file = Excel::crea(['Primo' => ['righe' => [['a']]], 'Secondo/con *segni?' => ['intestazioni' => ['X'], 'righe' => [['b']]]]);
$wb = dentro((string) $file, 'xl/workbook.xml');
prova(str_contains($wb, 'sheetId="1"') && str_contains($wb, 'sheetId="2"'), 'due fogli nello stesso file');
$nomi = [];
preg_match_all('/<sheet name="([^"]*)"/', $wb, $nomi);
prova($nomi[1] === ['Primo', 'Secondo con  segni '], 'i caratteri che Excel non ammette nel nome del foglio (/ * ? [ ] :) sono sostituiti');
prova(!str_contains(dentro((string) $file, 'xl/worksheets/sheet1.xml'), '<autoFilter'), 'senza intestazioni non c\'e filtro');
unlink((string) $file);

$lungo = str_repeat('Dipartimento di prova ', 5);
$file = Excel::crea([$lungo => ['righe' => [['x']]]]);
prova(mb_strlen((string) preg_replace('/.*<sheet name="([^"]*)".*/s', '$1', dentro((string) $file, 'xl/workbook.xml'))) <= 31, 'il nome del foglio non supera i 31 caratteri di Excel');
unlink((string) $file);

$file = Excel::crea(['Vuoto' => []]);
prova($file !== null && str_contains(dentro((string) $file, 'xl/worksheets/sheet1.xml'), '<sheetData>'), 'un foglio senza dati non fa saltare niente');
unlink((string) $file);

sezione('Documenti .docx');

$corpo = Word::paragrafo('Il **Consiglio** approva', ['al' => 'both']) . Word::tabella(['Esame', 'CFU'], [['Zoologia', '9']]);
$file = Word::crea($corpo, ['piede' => 'Verbale']);
prova($file !== null && is_file($file), 'il documento si crea');
$doc = dentro((string) $file, 'word/document.xml');
prova(str_contains($doc, '<w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">Consiglio</w:t>'), '**grassetto** nel testo diventa grassetto di Word');
prova(str_contains($doc, '<w:jc w:val="both"/>'), 'allineamento giustificato');
prova(str_contains($doc, 'Esame') && str_contains($doc, 'Zoologia'), 'la tabella ha intestazione e righe');
prova(str_contains(dentro((string) $file, 'word/footer1.xml'), 'Verbale'), 'il piede di pagina c\'e');
prova(str_contains(dentro((string) $file, '[Content_Types].xml'), 'wordprocessingml.document.main+xml'), 'il pacchetto OOXML si dichiara come documento');
unlink((string) $file);

$doc = Word::runs("prima\nseconda");
prova(str_contains($doc, '<w:br/>'), 'l\'andata a capo diventa un\'interruzione di riga');
prova(str_contains(Word::runs('x', ['b' => true, 'u' => true, 'i' => true, 'sz' => 24]), '<w:u w:val="single"/>'), 'grassetto, corsivo, sottolineato e corpo');
prova(str_contains(Word::paragrafo('x & y'), '&amp;'), 'il testo del paragrafo e protetto');

echo "\n$ok prove riuscite, $ko fallite.\n";
exit($ko > 0 ? 1 : 0);
