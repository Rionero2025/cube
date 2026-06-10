# CUBE Management Contract — Streamlit Pro Universal

Versione Streamlit rifattibile per gestione contratti, clienti CRM, documenti, lavori, pagamenti, rate/acconti/saldi e fatture interne.

## Avvio

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Note lettura contratti
- PDF testuali: lettura diretta.
- PDF grafici/moduli: lettura per blocchi + parser dedicato.
- Foto/scansioni: richiede Tesseract OCR installato sul PC e `pytesseract`.
- Tutti i dati importati restano modificabili prima e dopo il salvataggio.

## Sezioni principali
- Dashboard con Totale contratti, IVA contratti, Totale IVA inclusa, Incasso medio mensile.
- Importa contratto nel gestionale.
- Clienti CRM.
- Contratti modificabili.
- Pagamenti con celle editabili, acconti/saldi, allegati, modifica/elimina movimenti.
- Lavori cliente con allegati.
- Documenti.
- Fatture interne con modifica/elimina archivio.
