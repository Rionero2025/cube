# Fix pagine cliccabili e sito più veloce

Il pacchetto precedente conteneva le pagine HTML, ma non erano collegate al router pubblico del portale Streamlit.

Questa versione corregge:

- Menu e link cliccabili.
- Route pubbliche:
  - /?public_page=home
  - /?public_page=features
  - /?public_page=how
  - /?public_page=pricing
  - /?public_page=about
  - /?public_page=faq
  - /?public_page=contact
  - /?public_page=support
  - /?public_page=security
  - /?public_page=privacy
  - /?public_page=terms
  - /?public_page=cookie

- Le pagine HTML vengono renderizzate come statiche in iframe isolato, quindi sono più veloci e non vengono trasformate in testo da Streamlit.
- Login, registrazione e piani restano moduli Streamlit operativi.

Carica tutto lo ZIP su GitHub, fai commit e poi Manual Sync su Render.
