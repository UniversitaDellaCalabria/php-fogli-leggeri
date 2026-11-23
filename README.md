# php-fogli-leggeri

File **.xlsx** e **.docx** scritti in PHP puro: 260 righe in tutto, nessun pacchetto esterno, serve solo
l'estensione `zip`. Niente PhpSpreadsheet (che da sola pesa più di tutto il resto del tuo programma) e
niente PHPWord.

Esiste per un caso preciso e molto comune negli enti pubblici: **un server dove Composer non c'è**, e un
utente che ti chiede il foglio di calcolo da scaricare.

Requisiti: **PHP 8.1+**, estensioni `zip` e `mbstring`.

## Installazione

Con Composer:

```json
{
    "require": { "dodaroe/php-fogli-leggeri": "*" }
}
```

Senza Composer:

```php
require '/percorso/php-fogli-leggeri/autoload.php';
```

## Un foglio di calcolo

```php
use Dodaroe\FogliLeggeri\Download;
use Dodaroe\FogliLeggeri\Excel;

$file = Excel::crea([
    'Iscritti' => [
        'intestazioni' => ['Cognome', 'Nome', 'Email'],
        'righe' => [
            ['Rossi', 'Mario', 'mario@example.it'],
            ['Verdi', 'Anna', 'anna@example.it'],
        ],
        'larghezze' => [20, 20, 35],   // facoltativo: altrimenti le calcola
    ],
]);

if ($file === null) {
    exit('Manca l\'estensione zip di PHP.');
}

Download::invia($file, 'iscritti.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
```

L'intestazione esce in grassetto su fondo chiaro, con il **filtro** sulle colonne e la **prima riga
bloccata**: sono le tre cose che chi riceve un foglio si aspetta. Passa più chiavi all'array per avere più
fogli nello stesso file; i caratteri che Excel non ammette nei nomi dei fogli (`/ * ? [ ] :`) e il limite di
31 caratteri sono gestiti.

`Download::invia()` manda il file al browser, lo cancella e termina la richiesta. Se preferisci gestire tu la
risposta, ignoralo: `crea()` ti dà il percorso di un file temporaneo e il resto è tuo.

## Un documento

```php
use Dodaroe\FogliLeggeri\Word;

$corpo = Word::paragrafo('Verbale della seduta', ['b' => true, 'sz' => 28, 'al' => 'center'])
       . Word::paragrafo('Il **Consiglio** approva all\'unanimita.', ['al' => 'both'])
       . Word::tabella(['Esame', 'CFU'], [['Zoologia', '9'], ['Ecologia', '6']]);

$file = Word::crea($corpo, [
    'font' => 'Times New Roman',
    'logo' => __DIR__ . '/logo.png',   // nell'intestazione
    'testata' => 'Dipartimento di Prova',
    'piede_sx' => 'Il Segretario',
    'piede_dx' => 'Il Direttore',
]);
```

`**grassetto**` nel testo diventa grassetto di Word, e l'andata a capo diventa un'interruzione di riga. Il
numero di pagina nel piede c'è già.

## Prove

```bash
php tests/esegui.php
```

24 prove: aprono davvero i file prodotti e leggono l'XML dentro.

## Cosa questo pacchetto non fa

- **Non legge** file .xlsx o .docx: li scrive. Per leggerli serve altro.
- **Niente formule, grafici, immagini nelle celle, formati numerici, date, celle unite, colori per riga.**
  Le celle sono testo (`inlineStr`): un numero in un foglio prodotto così resta allineato come testo. Se ti
  serve un foglio di calcolo vero e proprio, con i numeri che si sommano, usa PhpSpreadsheet: questo
  pacchetto serve a consegnare un elenco.
- **Non impagina** il .docx: niente sezioni multiple, colonne, indice, stili personalizzati, note.
- **Non genera PDF.**
- **L'assistenza non è garantita.** È codice usato in produzione da un portale universitario e pubblicato
  perché possa servire ad altri, non un prodotto con un supporto dietro.

## Licenza

**AGPL-3.0-or-later**: vedi il file [LICENSE](LICENSE). Vuol dire che se usi questo codice in un programma
che offri ad altri, anche solo via rete, quel programma va pubblicato con la stessa licenza.
