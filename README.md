# Controllo consegne Nike

Interfaccia web per confrontare la **conferma d'ordine stagionale** con tutti i **file delle consegne** ricevuti durante la stagione.

I file vengono letti solo nel browser: non vengono caricati su nessun server. Restano memorizzati nel browser di quel computer finché non li rimuovi, così la volta successiva basta aggiungere le nuove consegne.

## Come si usa

1. Apri la pagina (vedi sotto).
2. **Conferma ordine stagionale**: trascina il file di conferma (es. `Happy Sport - Nike FA26 - Confirmation.xlsx`).
3. **File delle consegne**: trascina uno o più file di consegna. Puoi aggiungerne altri in qualsiasi momento.
4. Guarda i risultati e usa **Esporta Excel** o **Stampa / PDF** (nella finestra di stampa scegli "Salva come PDF").

## Cosa calcola

L'abbinamento è per **EAN + Stagione**, a prezzo **FPC netto IVA** dell'ordine.

| Scheda | Contenuto |
|---|---|
| Riepilogo | Ordinato, consegnato, residuo ed eccedenza in pezzi e valore, per stagione, reparto, numero d'ordine e genere |
| Consegne | Ogni consegna (file, packing list, truck, ordine): pezzi, valore FPC e valore netto fatturato |
| Eccedenze | Taglie consegnate in quantità superiore all'ordinato |
| Non ordinato | EAN consegnati che non sono nella conferma |
| Altra stagione | Articoli dell'ordine arrivati con un'altra stagione (non contano come consegnati) |
| Differenze prezzo/ordine | Prezzo FPC o numero d'ordine della consegna diverso da quello della conferma |
| Per articolo | Residuo per articolo (modello-colore) |
| Dettaglio EAN | Ogni taglia con stato: Da consegnare, Parziale, Completo, Eccedenza |

### Stagione

I codici stagione vengono uniformati: `FA26`, `FA 2026` e il codice numerico Nike `263` (anno 26 + stagione 3) diventano tutti `FA26` (1 = SP, 2 = SU, 3 = FA, 4 = HO). Se il codice non è riconoscibile, la stagione viene presa dal nome del truck (es. `BGRO093/FA26`). Sotto ogni file è indicata la conversione applicata.

### Colonne

Le colonne vengono riconosciute dal nome dell'intestazione (anche se non è nella prima riga). Sotto ogni file, in "Colonne riconosciute", puoi vedere e correggere quale colonna è usata per ogni dato; la scelta viene ricordata.

- **Conferma**: EAN UPC Cd, Qty status NNT, FPC Price IN EUR w/o VAT, Season, PO name SAP, BU, Article Number, Article Name, Color Desc, Size US, EU Size, Gender.
- **Consegne**: EAN, quantità, stagione, prezzo FPC, valore o prezzo netto, Po name, Truck, Packing list, articolo, descrizione, taglia.

## Pubblicazione

È una pagina statica (HTML + JavaScript, nessuna installazione):

- **In locale**: apri `index.html` con il browser (funziona anche offline, la libreria Excel è inclusa in `vendor/`).
- **GitHub Pages**: nel repository, Settings › Pages › "Deploy from a branch", ramo `main`, cartella `/ (root)`. La pagina sarà su `https://blstc11.github.io/controllo_consegne_nike/`.

## Struttura

- `index.html`, `css/style.css`: interfaccia
- `js/engine.js`: lettura dei file, riconoscimento colonne e confronto
- `js/app.js`: caricamento, tabelle, export Excel, stampa
- `vendor/xlsx.full.min.js`: [SheetJS](https://sheetjs.com) 0.18.5 (licenza Apache 2.0, `vendor/xlsx.LICENSE`)
