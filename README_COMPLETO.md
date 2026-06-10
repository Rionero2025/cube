# CUBE Management Contract — Streamlit Pro Completo

Versione finale completa con tutti i fix richiesti.

## Moduli inclusi
- Dashboard generale con riepiloghi economici contratti.
- Importa contratto da PDF/foto con lettura dati e importazione nel gestionale.
- Aziende/emittenti modificabili.
- Clienti CRM con scheda dati editabile.
- Contratti generabili e modificabili.
- Scelta template da usare per generare il contratto.
- Template contratti modificabili, creabili da testo incollato, Word o PDF.
- Pagamenti completi per tutti i clienti o per singolo cliente/contratto.
- Rate editabili, acconti, saldi, allegati pagamento, storico movimenti, modifica/elimina movimenti.
- Lavori CRM per cliente con allegati.
- Filtro lavori per intervallo date, tipo e stato.
- Generazione report PDF lavori con testo descrittivo automatico di riepilogo.
- Archivio documenti collegati a cliente/contratto/lavori.
- Fatture interne V1, PDF di cortesia, archivio, modifica/elimina, rigenerazione PDF.
- Impostazioni, cartelle allegati e database SQLite locale.

## Avvio
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Note importanti
- Per leggere foto/scansioni serve Tesseract OCR installato sul PC.
- Le fatture sono interne/di cortesia: non sono ancora XML SdI.
- Il database locale si trova in `data/cube_contracts_pro.db`.
- Gli allegati vengono archiviati nelle cartelle `allegati/`.
