# XLSX — rapporto di verifica

Data: 26 ottobre 2026. Repository di lavoro:
`illancillotto/rizzo-pii-cbo-version`, base `9057eef`, branch `feat/xlsx-support`.
`main` conservato sulla base, nessun merge/deploy, nessun fork remoto creato.
Tutte le fixture sono sintetiche. I pesi sono stati scaricati separatamente;
inferenza e prove sui documenti eseguite con modalità Hugging Face offline.

## Implementazione

- `src/app/pii_session.py`: identità esatta per categoria/tipo/valore,
  namespace casuale anticollisione, sessione condivisa su tutta la cartella.
- `src/app/xlsx_io.py`: preflight ZIP/XML/relazioni, lista di componenti
  supportati, budget, riconoscimento per cella e contesto intestazioni,
  colonne obbligatorie, revisione, export verificato e ripristino tipizzato.
- `src/app/xlsx_api.py`: nuove API locali, worker in background, anteprime e
  inventari paginati, owner cookie, same-origin/header/Host, scadenza,
  cancellazione, ammissione memoria, multipart senza spool su disco.
- `src/app/xlsx_ui.html`: flusso completo IT/EN, selettori colonne paginati,
  correzioni singole/associate, aggiunte, download separati e ripristino.
- `src/app/app.py`: adattatore batch allo stesso modello + regex + merge,
  registrazione API e collegamento Excel nell'interfaccia esistente.
- `tests/test_xlsx_io.py`, `tests/test_xlsx_api.py`: flusso deterministico e
  controlli negativi; non misurano qualità del modello.
- Manifest runtime/test, Dockerfile/server e builder, entrambi gli spec
  PyInstaller, script Linux/macOS e workflow Windows aggiornati; nuovo workflow
  `.github/workflows/xlsx-tests.yml`, senza dipendenze ML.
- Script riproducibili `scripts/check_xlsx_real_model.py`,
  `scripts/check_xlsx_browser.py`, `scripts/check_xlsx_libreoffice.py`.

## Baseline e regressioni

| Controllo | Passati | Falliti | Saltati | Note |
| --- | ---: | ---: | ---: | --- |
| Baseline iniziale, Python sistema | 73 | 0 | 10 | 83 totali, dipendenze app assenti |
| Baseline dopo installazione dipendenze | 83 | 0 | 0 | Pipeline simulata nei test che importano app |
| Suite finale `unittest discover -v tests` | 131 | 0 | 0 | 83 preesistenti + 48 XLSX/API |
| XLSX/API in venv isolato Python 3.11 | 48 | 0 | 0 | Solo requirements-xlsx-test.txt, senza torch |
| XLSX/API in venv isolato Python 3.14 | 48 | 0 | 0 | Solo requirements-xlsx-test.txt, senza torch |
| Ripristino JavaScript preesistente | 14 | 0 | 0 | `tests/test_ripristino_placeholder.py`, Node |
| Confini parola preesistente | 5 documenti | 0 | 0 | Backend con modello reale, fixture sintetiche |

La vecchia CI `tests.yml` mantiene il job standard-library e i suoi salti per
dipendenze mancanti. Il nuovo job XLSX installa esplicitamente Flask, openpyxl e
defusedxml e li importa prima dei test; esegue realmente la suite XLSX su Python
3.11/3.14 e il controllo JavaScript, senza installare torch. Si attiva sul push del
branch e sulle PR interessate. Nessun job leggero viene presentato come test del
modello. Le esecuzioni GitHub effettive vanno controllate sul commit pubblicato.

## Prove complete

**Modello reale, API — PASS.** `scripts/check_xlsx_real_model.py` importa l'app
senza sostituzioni della pipeline. Controlla cinque categorie: FULLNAME, CF,
PIVA, IBAN, EMAIL. Esegue upload → configurazione colonna numerica obbligatoria →
analisi → revisione singola/associata → export → download separati → riapertura →
ricerca dei cinque valori nelle parti ZIP → ripristino. Confronta 13 celle con
tipi, metadati, autore del commento, header di stampa, foglio veryHidden e riga
nascosta. Conserva una SUM su valori non sostituiti e verifica la cache vuota.
È un controllo di integrazione sintetico, non un benchmark del modello.

**Browser reale — PASS.** Chromium headless con profilo temporaneo, server Flask
locale e modello reale, Playwright. Upload, colonna obbligatoria, cambio lingua
senza perdere configurazione, revisione categoria, aggiunta manuale, export,
entrambi i download, upload del mapping e ripristino di sei celle/tipi,
cancellazione sessione. Zero errori JavaScript; richieste browser a host esterni
bloccate. Testa il browser web, non la WebView dell'installer Tauri.

**LibreOffice headless — PASS.** Versione 26.2.6.3, fixture esclusivamente
sintetica con rilevatore regex. Apre/risalva il file esportato, conserva i
placeholder e ricalcola `SUM(B2:B3)` al valore atteso. La copia risalvata è
correttamente rifiutata dal mapping originale per SHA diverso. Non misura la
qualità del modello né certifica tutte le funzioni di Calc/Excel.

**Controlli negativi automatici — PASS.** Sessioni distinte, collisioni,
namespace già presente, varianti esatte e tipi, zeri iniziali, righe/colonne/fogli
nascosti, celle unite, Unicode e multilinea, commenti a run separati,
proprietà/header/footer, nomi, hyperlink, formule vietate e dipendenze,
cache, componenti attivi/complessi, file corrotti e legacy/cifrati, DTD/XML
sconosciuti, ZIP duplicati/commenti, budget e area unioni, residui noti,
alterazioni strutturali, no-mapping, dizionari errati/incompleti/duplicati e
rielaborazione con mapping distinti, API, scadenza, cancellazione e isolamento
con ID noto da un altro browser.

## Ambiente e distribuzione

Ambiente principale: Linux x86_64, Python 3.14.4, Flask 3.1.3, openpyxl 3.1.5,
defusedxml 0.7.1, torch 2.14.1+cpu, transformers 4.57.3, PyMuPDF 2.0,
modello `rizzoaiacademy/rizzo-pii-0.3B` revisione `v1.5.0`.
Anche Python 3.11 verificato separatamente per suite leggera.

- JavaScript Excel: sintassi Node verificata e percorso eseguito nel browser.
- Script shell: `bash -n` sui tre script di build — PASS.
- Spec PyInstaller: parsing Python dei due spec — PASS.
- Compose: `docker compose config --quiet` — PASS.
- YAML workflow: parsing dei job — PASS.
- Installer Windows/macOS/Linux, sidecar PyInstaller e build dell'immagine Docker:
  **non costruiti né collaudati** in questa sessione. I file di distribuzione
  sono aggiornati, ma non c'è una dichiarazione di installer verificato.
- Microsoft Excel desktop e Tauri WebView: **non provati**.

Il caricamento del modello emette il warning già presente di transformers sul
regex del tokenizer (compatibilità Mistral) e PyMuPDF segnala la deprecazione
dell'import `fitz`. Non sono stati alterati tokenizer, modello o pipeline
preesistente per eliminare questi avvisi; i controlli del modello sopra passano,
ma non sostituiscono una valutazione di qualità estesa.

## Limiti residui e operatività

Nessun blocco impedisce il flusso locale sul sottoinsieme supportato. Restano
le restrizioni intenzionali della v1 dettagliate in [XLSX.md](XLSX.md): componenti
complessi rifiutati, solo tema Office standard, formule limitate, perdita dei run
di stile nei testi modificati, nessuna ricostruzione byte per byte e mapping
utilizzabile soltanto sul download con hash corrispondente.

Il tempo è controllato fra batch, non come interruzione nativa a metà inferenza;
il budget memoria è stimato, da affiancare a cgroup/Docker per un tetto rigido.
I falsi negativi del riconoscimento rimangono possibili. Non c'è fallback che
consenta il download ordinario dopo un controllo strutturale/residui fallito.

Per avviare l'ambiente preparato:

```bash
PII_MODEL_DIR="$PWD/models/rizzo-pii-0.3B-v1.5.0" \
HF_HUB_OFFLINE=1 TRANSFORMERS_OFFLINE=1 OMP_NUM_THREADS=4 \
.venv/bin/python src/app/serve.py --host 127.0.0.1 --port 5005
# aprire http://127.0.0.1:5005/xlsx
```

Piano/baseline: [XLSX_PLAN.md](XLSX_PLAN.md). I commit separano piano, motore,
integrazione, hardening/prove e distribuzione/documentazione; l'identificativo
finale e lo SHA remoto sono riportati nella consegna e verificabili con
`git log --oneline main..feat/xlsx-support` e
`git ls-remote origin refs/heads/feat/xlsx-support`.
