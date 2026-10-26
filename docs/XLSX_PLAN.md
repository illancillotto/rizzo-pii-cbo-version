# Supporto XLSX — piano e baseline

Base: fork `illancillotto/rizzo-pii-cbo-version`, commit `9057eef`.
Branch: `feat/xlsx-support`. Nessuna modifica preesistente.

## Baseline (26 ottobre 2026)

`python3 -m unittest discover -v tests`: 83 test, 73 passati,
0 falliti, 10 saltati (dipendenze Flask/PyMuPDF/torch assenti).
Nessun peso del modello presente. Il test manuale JS non è scoperto da unittest.
La CI attuale non installa dipendenze: salta i test dell'app e non verifica il
modello. I test con pipeline simulata non misurano la qualità del modello.

## Analisi

`app.py` carica la pipeline token-classification all'import; inferenza per chunk
con overlap. `_merge` privilegia checksum e regex, poi score e lunghezza.
`analyze` ricrea contatori e mapping per chiamata e normalizza spazi/maiuscole:
non può quindi essere usato cella per cella per un ripristino Excel esatto.
Preferenze in `server_config.py`, ripristino testo in JavaScript; PDF in
`pdf_export.py`. Upload globale 50 MiB; documenti PDF in LRU in memoria.
Non risultano protezioni Origin/CSRF esplicite: le nuove API avranno controllo
same-origin, token casuale per documento e cache disabilitata.
Runtime, Docker, entrambi gli spec PyInstaller, script Linux/macOS/Windows e
workflow Windows devono includere le nuove dipendenze. Nessuna build installer
verrà dichiarata verificata senza costruirla e provarla.

## Fasi / commit previsti

1. Questo piano e baseline.
2. Sessione indipendente da Flask e lettura/scrittura OOXML conservativa,
   preflight, limiti, formule, mapping tipizzato e verifica residui.
3. API isolate e interfaccia IT/EN per configurazione/revisione/ripristino.
4. Test sintetici, regressioni, CI leggera, distribuzione e guida.

Nessun riaddestramento o servizio cloud. Il pacchetto viene modificato in memoria
senza estrazione su disco; le parti non gestibili vengono rifiutate. Gli stili
si conservano nel pacchetto originale, non attraverso una riscrittura generale
di openpyxl; openpyxl serve come lettore indipendente di verifica.

Aggiornamenti e risultati finali in `docs/XLSX_VERIFICATION.md`.
