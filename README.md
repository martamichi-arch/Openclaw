# StudyBook Builder

App web minimale per caricare appunti, slide, testi di riferimento e generare un unico documento Word di studio.

## Funzioni

- Caricamento PDF, DOCX, PPTX, TXT
- Estrazione testo
- Estrazione immagini originali dai file
- Generazione documento Word
- Corpo in Aptos 10
- Titoli e sottotitoli formattati
- Uso opzionale di OpenAI per ricostruire spiegazioni più leggibili

## Cartelle logiche

- Appunti studenti
- Slide professore
- Testi di riferimento
- Appunti / extra

## Secrets Streamlit consigliati

Nel pannello Streamlit Cloud aggiungere:

```toml
APP_PASSWORD = "abcdefghi"
OPENAI_API_KEY = "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"
