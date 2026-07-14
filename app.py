import io
import os
import tempfile
from datetime import datetime
from pathlib import Path

import fitz
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

st.set_page_config(
    page_title="StudyBook Builder",
    page_icon="📚",
    layout="wide",
)

def get_secret(name, default=None):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return os.getenv(name, default)

APP_PASSWORD = ***"APP_PASSWORD", "")

def check_password():
    if not APP_PASSWORD:
        *** True

    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    st.title("📚 StudyBook Builder")
    password = st.text_input("Password", type="password")

    if st.button("Entra"):
        if password == APP_PASSWORD:
            st.ses…ated = True
            st.rerun()
        else:
            st.error("Password errata.")

    return False

if not check_password():
    st.stop()

st.title("📚 StudyBook Builder")
st.caption("Genera un unico documento Word da appunti, slide e testi di riferimento.")

def clean_text(text):
    if not text:
        return ""
    return "\n".join(line.strip() for line in text.splitlines() if line.strip())

def extract_pdf(uploaded_file, category):
    result = {
        "filename": uploaded_file.name,
        "category": category,
        "text_blocks": [],
        "images": [],
    }

    data = uploaded_file.read()
    doc = fitz.open(stream=data, filetype="pdf")

    for page_index, page in enumerate(doc):
        text = clean_text(page.get_text("text"))

        if text:
            result["text_blocks"].append({
                "source": f"{uploaded_file.name} — pagina {page_index + 1}",
                "text": text,
            })

        for image_index, img in enumerate(page.get_images(full=True)):
            try:
                xref = img[0]
                base = doc.extract_image(xref)
                image_bytes = base["image"]
                ext = base.get("ext", "png")

                result["images"].append({
                    "source": f"{uploaded_file.name} — pagina {page_index + 1}",
                    "filename": f"{Path(uploaded_file.name).stem}_p{page_index + 1}_{image_index + 1}.{ext}",
                    "bytes": image_bytes,
                    "context": text[:1200],
                })
            except Exception:
                pass

    return result

def extract_docx(uploaded_file, category):
    result = {
        "filename": uploaded_file.name,
        "category": category,
        "text_blocks": [],
        "images": [],
    }

    data = uploaded_file.read()
    document = Document(io.BytesIO(data))

    paragraphs = []

    for p in document.paragraphs:
        t = clean_text(p.text)
        if t:
            paragraphs.append(t)

    full_text = "\n".join(paragraphs)

    if full_text:
        result["text_blocks"].append({
            "source": uploaded_file.name,
            "text": full_text,
        })

    rels = document.part._rels
    count = 0

    for rel in rels:
        rel_obj = rels[rel]

        if "image" in rel_obj.target_ref:
            try:
                count += 1
                image_bytes = rel_obj.target_part.blob

                result["images"].append({
                    "source": uploaded_file.name,
                    "filename": f"{Path(uploaded_file.name).stem}_image_{count}.png",
                    "bytes": image_bytes,
                    "context": full_text[:1200],
                })
            except Exception:
                pass

    return result

def extract_pptx(uploaded_file, category):
    result = {
        "filename": uploaded_file.name,
        "category": category,
        "text_blocks": [],
        "images": [],
    }

    data = uploaded_file.read()
    prs = Presentation(io.BytesIO(data))

    for slide_index, slide in enumerate(prs.slides):
        slide_texts = []

        for shape in slide.shapes:
            if hasattr(shape, "text"):
                t = clean_text(shape.text)
                if t:
                    slide_texts.append(t)

        slide_text = "\n".join(slide_texts)

        if slide_text:
            result["text_blocks"].append({
                "source": f"{uploaded_file.name} — slide {slide_index + 1}",
                "text": slide_text,
            })

        image_count = 0

        for shape in slide.shapes:
            if getattr(shape, "shape_type", None) == 13:
                try:
                    image_count += 1
                    image = shape.image
                    image_bytes = image.blob
                    ext = image.ext or "png"

                    result["images"].append({
                        "source": f"{uploaded_file.name} — slide {slide_index + 1}",
                        "filename": f"{Path(uploaded_file.name).stem}_slide_{slide_index + 1}_{image_count}.{ext}",
                        "bytes": image_bytes,
                        "context": slide_text[:1200],
                    })
                except Exception:
                    pass

    return result

def extract_txt(uploaded_file, category):
    result = {
        "filename": uploaded_file.name,
        "category": category,
        "text_blocks": [],
        "images": [],
    }

    raw = uploaded_file.read()

    try:
        text = raw.decode("utf-8")
    except Exception:
        text = raw.decode("latin-1", errors="ignore")

    text = clean_text(text)

    if text:
        result["text_blocks"].append({
            "source": uploaded_file.name,
            "text": text,
        })

    return result

def extract_file(uploaded_file, category):
    suffix = Path(uploaded_file.name).suffix.lower()

    if suffix == ".pdf":
        return extract_pdf(uploaded_file, category)

    if suffix == ".docx":
        return extract_docx(uploaded_file, category)

    if suffix == ".pptx":
        return extract_pptx(uploaded_file, category)

    if suffix == ".txt":
        return extract_txt(uploaded_file, category)

    return {
        "filename": uploaded_file.name,
        "category": category,
        "text_blocks": [],
        "images": [],
    }

def split_text(text, max_chars=9000):
    chunks = []
    current = []
    total = 0

    for paragraph in text.split("\n"):
        if total + len(paragraph) > max_chars and current:
            chunks.append("\n".join(current))
            current = []
            total = 0

        current.append(paragraph)
        total += len(paragraph)

    if current:
        chunks.append("\n".join(current))

    return chunks

def ai_rewrite(all_blocks, correction_notes):
    api_key = get_secret("OPENAI_API_KEY", "")

    if not api_key or OpenAI is None:
        return None

    client = OpenAI(api_key=api_key)

    ordered_text = []

    for block in all_blocks:
        ordered_text.append(
            f"[FONTE: {block['source']}]\n{block['text']}"
        )

    full_text = "\n\n".join(ordered_text)
    chunks = split_text(full_text)

    outputs = []

    for i, chunk in enumerate(chunks, start=1):
        prompt = f"""
Devi trasformare questi materiali universitari in un testo unico di studio, in italiano.

Regole obbligatorie:
- Non omettere informazioni.
- Non inventare.
- Non generare immagini.
- Usa come base principale gli appunti studenti.
- Integra in seconda battuta con le slide del professore.
- Usa i testi di riferimento per chiarire, completare e rendere comprensibile.
- Mantieni un tono da libro di testo.
- Organizza in capitoli, sezioni e sottosezioni.
- Se una fonte è confusa, ricostruisci la spiegazione ma senza perdere i dettagli.
- Conserva concetti, definizioni, elenchi, esempi, dat
...(truncated)...
