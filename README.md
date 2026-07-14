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
APP_PASSWORD = "scegli…word"
OPENAI_API_KEY = "la-tua…enai"
