# Excel .xlsx — guida e contratto di supporto v1

Rizzo PII elabora cartelle Excel in locale usando lo stesso modello e la stessa
rete regex/checksum di testo e PDF. La nuova pagina **Excel .xlsx** è raggiungibile
dall'app o direttamente su `/xlsx`. Non usa API cloud, non valuta formule e non
scarica risorse nel browser. I pesi del modello devono essere già disponibili.

Questa è **pseudonimizzazione**, quando esiste un dizionario reversibile. La
modalità senza dizionario non crea una chiave esportabile, ma non certifica
l'anonimato: contesto, dati non rilevati e collegabilità possono identificare una
persona. Un controllo senza residui noti non garantisce l'assenza di ogni PII.

## Uso

1. Carica una copia `.xlsx` non cifrata. Il controllo del pacchetto precede
   l'analisi. Se incontra una parte non supportata, spiega il motivo e blocca il
   documento; non produce una versione apparentemente ripulita.
2. Controlla fogli, dimensioni e stato `visible`, `hidden`, `veryHidden`. Tutti i
   fogli sono selezionati inizialmente. Un foglio deselezionato viene comunque
   controllato: PII rilevata e non esclusa blocca il documento.
3. Imposta per ogni foglio riga intestazione (`0` = assente) e prima riga dati.
   Le intestazioni non sono sicure per definizione: passano dal riconoscimento.
4. Seleziona le colonne obbligatorie e assegna categoria/prefisso (`A-Z`, cifre,
   underscore, massimo 32 caratteri). Ogni valore non vuoto nella riga dati viene
   sostituito per intero, inclusi numeri, booleani e date. La riga intestazione è
   esclusa dalla regola obbligatoria, ma non dal controllo automatico.
5. Seleziona le categorie da lasciare visibili, per esempio `DATE` e `AMOUNT`.
   Le colonne obbligatorie **prevalgono** su queste esclusioni; l'anteprima indica
   il conflitto. Impostazioni iniziali e stato del mapping provengono dalle
   preferenze esistenti, gli override Excel sono limitati alla sessione.
6. Avvia l'analisi. Il server mostra fase e campi completati, con elaborazione
   in background e batch limitati. Il modello non vede celle concatenate.
7. Rivedi le proposte, 50 per pagina. Cambia categoria o escludi un falso
   positivo: **Occorrenza** modifica solo quella proposta, **Associate** tutte le
   proposte che condividono lo stesso valore esatto, tipo e categoria.
   Una regola obbligatoria non può essere disabilitata in revisione: cambia la
   configurazione e analizza nuovamente se la regola era sbagliata.
8. Aggiungi dati mancanti con foglio/cella oppure con ID campo. L'inventario
   paginato permette di controllare anche metadati e commenti. Per cambiare
   l'intervallo di una proposta, escludila e aggiungi la nuova sostituzione.
   Gli offset sono indici di caratteri Unicode Python, inizio incluso e fine
   esclusa; il campo fine vuoto comprende tutto il valore. Testi lunghi sono
   troncati a 240 caratteri nell'anteprima, il valore completo resta sul server.
9. Premi **Verifica ed esporta**. Il download è disponibile solo dopo riapertura,
   verifica strutturale e controllo residui. Una correzione successiva invalida
   sia il download sia il dizionario precedente.
10. Scarica `.xlsx` e, se abilitato, **separatamente** il dizionario `.json`.
    Non inviare il dizionario insieme al documento pseudonimizzato.
11. Per ripristinare, carica il file esportato non modificato e il suo dizionario
    nella sezione dedicata. Conserva questi due file se vuoi ripristinare in
    seguito, anche dopo la scadenza della sessione.

Disattivando il riconoscimento automatico si disattivano le proposte automatiche
per le celle, ma restano controllo dei metadati e verifica finale. Il download
resta bloccato se il rilevatore trova PII non coperta dalle regole/revisioni.

## Identità, sessioni e mapping

Ogni cartella ha una sessione indipendente. Identità = **categoria + tipo OOXML +
valore esatto**, senza trim, casefold, normalizzazione Unicode o unificazione di
spazi. `Alfa`, `alfa`, ` Alfa`, il numero `42` e il testo `42` restano distinti.
Lo stesso valore nella stessa categoria/tipo usa lo stesso placeholder su tutti
i fogli. Le varianti di scrittura non vengono collegate automaticamente: è una
scelta necessaria per recuperare la forma originale.

Formato dei placeholder: `[[RZ_<namespace casuale>_<CATEGORIA>_<indice>]]`.
Il namespace viene rigenerato se compare nei campi iniziali. Placeholder già
presenti sono opachi e non vengono scambiati per nuove entità; quelli della
precedente elaborazione non sono ripristinati dal dizionario nuovo. L'indice
interno usa HMAC con chiave casuale in memoria; non viene esportato.

Senza mapping non vengono costruite voci placeholder→originale, né restituiti
valori originali in anteprima/inventario, né salvati dizionari nel browser.
Il documento sorgente rimane temporaneamente in memoria per revisione e controlli.
Né documenti né dizionari vengono salvati permanentemente dal server, inclusi gli
upload multipart che sono mantenuti in memoria senza spool su disco.

Il dizionario versione 1 ha `format: "rizzo-pii-xlsx"`, `version`, `namespace`,
`sha256` del **file esportato** ed `edits` tipizzate per campo. Contiene gli
originali dei soli campi modificati (per una cella di testo libero, il campo
intero) e i valori attesi dopo le modifiche. È un file sensibile non cifrato.
L'hash lega il mapping a un'esportazione, non è una firma contro modifiche
intenzionali del dizionario. Usare soltanto dizionari locali affidabili.

Il ripristino recupera valori e tipi supportati: numeri, stringhe con zeri
iniziali, date OOXML e seriali Excel col formato originale, booleani, errori,
metadati e commenti. **Non** ricostruisce i byte originali né i singoli run di
formattazione di testi/commenti modificati. Le cache formula restano svuotate.
Non viene promessa la ricostruzione di firme, CRC, compressione o ordine XML.

File modificato, placeholder modificato/sconosciuto, mapping incompleto,
duplicazioni di campo, formato/versione errati e dizionario di un'altra
esportazione vengono rifiutati. **Salvare di nuovo in Excel/LibreOffice cambia
l'hash**, quindi il mapping originale non ripristina quella copia. Conservare
il download originale prima di ricalcolarlo o modificarlo. Il ripristino produce
un nuovo `.xlsx`, non sovrascrive il file selezionato.

## Caratteristiche supportate

La v1 accetta solo OOXML SpreadsheetML transizionale con parti e relazioni nella
lista consentita. Gli XML vengono modificati direttamente senza estrazioni ZIP
sul filesystem. `openpyxl` riapre indipendentemente il risultato.

| Caratteristica | Comportamento verificato |
| --- | --- |
| Testo, Unicode, multilinea, celle vuote | Analisi per cella; vuoti intatti |
| Numeri, date, booleani, errori | Rilevamento e regole obbligatorie; ripristino tipizzato |
| Intestazioni di colonna | Riconosciute come testo; contesto per CF/PIVA/telefono/matricola/utenza, senza inserirlo nelle celle |
| Hidden/veryHidden, righe/colonne nascoste | Inclusi; stato conservato |
| Celle unite | Si elabora il valore memorizzato; intervalli conservati entro il budget |
| Stili, formati numerici, larghezze/altezze | XML conservato; test automatici su proprietà rappresentative |
| Rich text nelle celle | Contenuto unito per riconoscimento; se modificato diventa testo semplice, stile di cella conservato |
| Commenti tradizionali | Testo e autore riconosciuti; VML limitato alla geometria della nota; testo modificato senza run originali |
| Header/footer di stampa | Contenuto riconosciuto, codici di formattazione conservati, font standard; limite 255 unità UTF-16 |
| Proprietà core/app | Campi testuali riconosciuti; date tipizzate mascherate con data neutra valida e reversibile |
| Filtri semplici | Valori dei filtri riconosciuti e verificati |
| Nomi fogli, nomi definiti, nomi/font/formati degli stili | Controllati, senza rinomina automatica: PII rilevata blocca l'analisi |
| Nomi definiti | Solo intervalli statici su fogli locali; niente costanti o espressioni |
| Stringhe condivise | Espansione in celle; tabella sorgente rimossa esplicitamente, inclusi valori inutilizzati |
| Tema | Solo tema Office standard confrontato con quello distribuito da openpyxl |

Rifiutati: `.xls`, `.xlsb`, `.xlsm`, file cifrati o protetti; macro, controlli,
immagini/disegni, oggetti incorporati, pivot/cache, connessioni/query, collegamenti
esterni, hyperlink anche interni, tabelle strutturate, grafici, commenti threaded,
custom XML/proprietà custom, estensioni, formattazione condizionale, convalide,
schemi strict o sconosciuti, temi personalizzati, formule complesse, escape
testuali `_xNNNN_`, commenti/istruzioni XML e campi ZIP aggiuntivi.
La risposta indica la classe di problema senza riportare contenuti personali.
Un preflight riuscito mostra `unsupported: []`; un preflight fallito non conserva
il documento e non rende disponibile alcun download.

## Formule

Nessuna formula viene eseguita dal server. Sono accettati riferimenti A1 e range
**sullo stesso foglio**, costanti numeriche e operatori aritmetici/confronti;
funzioni consentite: `SUM`, `MIN`, `MAX`, `AVERAGE`, `COUNT`, `ABS`, `ROUND`.
La sintassi viene tokenizzata e verificata, non viene usato find-and-replace.

Bloccati stringhe letterali, concatenazioni, DDE, riferimenti esterni/tra fogli,
nomi in formula, formule condivise/matriciali/dinamiche e altre funzioni.
Costanti numeriche sospette passano anche dalla rete regex e dalla ricerca dei
residui noti. La v1 non certifica l'assenza di identificativi arbitrari presenti
nelle costanti: controllare il documento prima di condividerlo.

Se una formula riferisce una cella da sostituire, l'export viene bloccato: niente
risultati alterati o stringhe al posto di numeri senza avviso. Se una colonna
obbligatoria contiene formule, viene bloccata già l'analisi. Preparare
esplicitamente una copia con formule convertite in valori quando necessario.
Le formule preservate hanno cache risultati rimosse e ricalcolo completo
richiesto all'apertura. LibreOffice headless ha verificato un caso `SUM`; non è
una verifica completa di tutte le funzioni in ogni versione di Excel.

## Limiti e ciclo di vita

| Variabile ambiente | Default |
| --- | ---: |
| `PII_XLSX_COMPRESSED` | 10 MiB |
| `PII_XLSX_EXPANDED` | 50 MiB |
| `PII_XLSX_PARTS` | 500 parti ZIP |
| `PII_XLSX_SHEETS` | 32 fogli |
| `PII_XLSX_CELLS` | 50.000 celle e area complessiva dei fogli/unioni |
| `PII_XLSX_CELL_LENGTH` | 32.767 unità UTF-16, massimo fisico Excel |
| `PII_XLSX_CHARACTERS` | 2.000.000 caratteri complessivi |
| `PII_XLSX_XML_NODES` | 200.000 nodi XML complessivi |
| `PII_XLSX_SCAN_NODES` | 200.000 nodi nell'indice residui |
| `PII_XLSX_SECONDS` | 120 secondi per analisi/esportazione |
| `PII_XLSX_SESSIONS` | 6 documenti |
| `PII_XLSX_MEMORY` | 512 MiB, budget conservativo di ammissione |
| `PII_XLSX_TTL` | 900 secondi dal caricamento |

Profondità XML massima 64, massimo 64 attributi per nodo, formule massimo 2.048
caratteri, pagina anteprima massimo 100 proposte (default UI 50). Le dimensioni e
le stringhe condivise vengono controllate prima dell'espansione; l'area delle
unioni prima della riapertura normale openpyxl. Il limite di tempo è cooperativo
tra batch: non interrompe a metà una chiamata nativa del modello già in corso.
Il budget di memoria è un'ammissione stimata, non un limite RSS del processo;
per un tetto rigido usare anche il limite di memoria Docker/cgroup.

Identificativi documento casuali a 256 bit e cookie proprietario HttpOnly,
SameSite Strict separano i browser. Lo store e la coda risiedono nello stesso
processo, un worker d'inferenza Excel; configurare **un worker Gunicorn**, come
nel Dockerfile. Non c'è conservazione permanente implicita; la scadenza cancella
entro il passaggio dello sweeper (massimo 30 secondi), o al prossimo accesso.
Un'elaborazione cancellata trattiene la propria riserva fino all'uscita dal batch
per non aggirare i limiti tramite cancellazioni/caricamenti ripetuti.

Le nuove API richiedono Host loopback, stessa origine quando indicata,
`X-Rizzo-XLSX: 1` per mutazioni, e non espongono CORS. Per un'installazione
espressamente gestita su altro hostname usare `PII_XLSX_TRUSTED_HOSTS` (nomi
separati da virgola). Non è un sistema multiutente autenticato per Internet.
Risposte `no-store`, errori e log applicativi non contengono valori di celle,
nomi file caricati o dizionari. Gli access log possono riportare l'ID casuale;
da solo non autorizza l'accesso senza il cookie proprietario.

## API

Tutte le mutazioni: header `X-Rizzo-XLSX: 1`. Conservare il cookie ricevuto.

| Metodo e percorso | Input / risultato |
| --- | --- |
| `POST /api/xlsx/upload` | Multipart `file` → `201 {id, sheets, cells, fields, formulas, policy, ttl_seconds}` |
| `POST /api/xlsx/<id>/analyze` | JSON configurazione → `202`; sostituisce eventuale analisi precedente |
| `GET /api/xlsx/<id>/status` | Stato, progressi numerici, errore privo di contenuti |
| `GET /api/xlsx/<id>/preview?page=0&size=50` | Proposte e coordinate paginate |
| `GET /api/xlsx/<id>/fields?page=0&size=25&scope=metadata` | Inventario; scope `metadata`, `cells`, `all` |
| `PATCH /api/xlsx/<id>/review` | `{id, category?, enabled?, scope: "occurrence"\|"associated"}` |
| `POST /api/xlsx/<id>/review` | `{sheet, cell, category, start?, end?}` oppure `{slot, ...}` |
| `POST /api/xlsx/<id>/export` | `202`; verifica e prepara file e mapping separati |
| `GET /api/xlsx/<id>/download` | File soltanto in stato `ready` |
| `GET /api/xlsx/<id>/mapping` | Dizionario soltanto se abilitato e output verificato |
| `DELETE /api/xlsx/<id>` | Cancella/annulla la sessione |
| `POST /api/xlsx/restore` | Multipart `file` + `mapping` → nuovo `.xlsx` |

Esempio configurazione (indici fogli da 0, colonne e righe da 1):

```json
{
  "sheets": [0, 1],
  "header_rows": {"0": 1, "1": 3},
  "automatic": true,
  "include_mapping": true,
  "exclude_tags": ["AMOUNT", "DATE"],
  "columns": [
    {"sheet": 0, "column": 2, "header_row": 1, "start_row": 2, "category": "ACCOUNT"}
  ]
}
```

Il download ordinario resta bloccato per residui non autorizzati, errori o
parti non ispezionabili. Le sole autorizzazioni a mantenere dati rilevati sono
le categorie esplicitamente escluse e le proposte esplicitamente disabilitate,
limitate al campo e al valore revisionato.

## Verifiche riproducibili

```bash
# Flusso deterministico + regex; nessuna dipendenza ML
python -m pip install -r requirements-xlsx-test.txt
python -m unittest discover -v -s tests -p 'test_xlsx*.py'

# Regressioni dell'app (richiede dipendenze runtime, usa pipeline simulata)
python -m unittest discover -v tests
python tests/test_ripristino_placeholder.py
python tests/test_confini_parola.py

# Riconoscimento + ciclo API con modello reale già locale; nessun download
PII_MODEL_DIR="$PWD/models/rizzo-pii-0.3B-v1.5.0" python scripts/check_xlsx_real_model.py

# Browser Chromium di test, profilo temporaneo; avviare prima il server locale
python -m pip install playwright
python scripts/check_xlsx_browser.py

# Interoperabilità, fixture sintetica regex (richiede LibreOffice)
python scripts/check_xlsx_libreoffice.py
```

Risultati e limiti della compatibilità effettivamente provata:
[XLSX_VERIFICATION.md](XLSX_VERIFICATION.md).
