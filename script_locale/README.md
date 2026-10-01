# Script locale (macOS)

Stesso controllo dell'interfaccia web, ma eseguito sul computer a partire da una cartella.

## Cartelle

Lo script si aspetta, accanto a sé:

- `ordini_confermati/`: il file di conferma della stagione (es. `Happy Sport - Nike FA26 - Confirmation.xlsx`)
- `consegne/`: tutti i file di consegna ricevuti (es. `Happy Sport - Delivery File - 01.09.2026.xlsx`)
- `output/`: qui vengono scritti `controllo_consegne_<stagione>.xlsx` e `controllo_consegne_<stagione>.pdf`

## Uso

Doppio clic su **Aggiorna controllo consegne.command**: rigenera Excel e PDF e apre la cartella `output`.
In alternativa, da Terminale: `python3 controllo_consegne.py`.

## Requisiti

- Python 3 con `pandas` e `openpyxl` (`python3 -m pip install pandas openpyxl`)
- Google Chrome in `/Applications/Google Chrome.app`: viene usato in modalità headless per creare il PDF

## Stagione

Nei file di consegna la stagione è un codice numerico. La tabella `SEASON_CODES` in `controllo_consegne.py` lo converte (263 = FA26); se un codice manca, la stagione viene presa dal suffisso del Truck (es. `BGRO093/FA26`). Per una nuova stagione aggiungi il codice alla tabella.
