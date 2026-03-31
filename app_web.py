import streamlit as st
import pdfplumber
import re
from docx import Document
from docx.shared import Pt
import os

# IA (opcional)
try:
    from openai import OpenAI
    client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
    IA_ATIVA = True
except:
    IA_ATIVA = False

st.set_page_config(page_title="Comparador de Portarias")

st.title("📄 Comparador Profissional de Portarias")
st.write("Envie dois PDFs e gere automaticamente a versão comparada.")

# =========================
# FUNÇÕES
# =========================
def extrair_texto(pdf):
    texto = ""
    with pdfplumber.open(pdf) as p:
        for page in p.pages:
            texto += (page.extract_text() or "") + "\n"
    return texto


def extrair_blocos(texto):
    padrao = r"(Art\. ?\d+º?|§ ?\d+º?|[0-9]+\.[0-9\.]+)"
    partes = re.split(padrao, texto)
    blocos = {}

    for i in range(1, len(partes), 2):
        chave = partes[i]
        conteudo = partes[i+1] if i+1 < len(partes) else ""
        blocos[chave] = conteudo

    return blocos


def comparar(b1, b2):
    resultado = []
    chaves = set(b1.keys()).union(b2.keys())

    for c in sorted(chaves):
        t1 = b1.get(c, "")
        t2 = b2.get(c, "")

        if t1 == t2:
            resultado.append(("igual", c, t1))
        elif t1 and not t2:
            resultado.append(("removido", c, t1))
        elif not t1 and t2:
            resultado.append(("adicionado", c, t2))
        else:
            resultado.append(("alterado", c, t1, t2))

    return resultado


def gerar_docx(resultado):
    doc = Document()

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    for item in resultado:
        tipo = item[0]

        if tipo == "igual":
            doc.add_paragraph(f"{item[1]} {item[2]}")

        elif tipo == "removido":
            p = doc.add_paragraph()
            r = p.add_run(f"{item[1]} {item[2]}")
            r.font.strike = True

        elif tipo == "adicionado":
            p = doc.add_paragraph()
            r = p.add_run(f"{item[1]} {item[2]}")
            r.bold = True

        elif tipo == "alterado":
            p1 = doc.add_paragraph()
            r1 = p1.add_run(f"{item[1]} {item[2]}")
            r1.font.strike = True

            p2 = doc.add_paragraph()
            r2 = p2.add_run(f"{item[1]} {item[3]}")
            r2.bold = True

    caminho = "Portaria_Comparada.docx"
    doc.save(caminho)
    return caminho


def comparar_ia(t1, t2):
    if not IA_ATIVA:
        return "❌ IA não configurada. Adicione sua chave da OpenAI."

    prompt = f"""
Compare juridicamente os textos abaixo:

- Texto antigo riscado
- Texto novo abaixo
- Manter estrutura

TEXTO 1:
{t1}

TEXTO 2:
{t2}
"""

    resp = client.chat.completions.create(
        model="gpt-5",
        messages=[{"role": "user", "content": prompt}]
    )

    return resp.choices[0].message.content


# =========================
# INTERFACE
# =========================
pdf1 = st.file_uploader("📥 Arraste o PDF ORIGINAL", type="pdf")
pdf2 = st.file_uploader("📥 Arraste o PDF ALTERADO", type="pdf")

modo = st.radio("Modo:", ["Normal", "IA (mais preciso)"])

if st.button("🚀 Gerar comparação"):

    if pdf1 and pdf2:

        t1 = extrair_texto(pdf1)
        t2 = extrair_texto(pdf2)

        if modo == "Normal":
            b1 = extrair_blocos(t1)
            b2 = extrair_blocos(t2)
            resultado = comparar(b1, b2)

            arquivo = gerar_docx(resultado)

            with open(arquivo, "rb") as f:
                st.download_button("📄 Baixar Word", f, file_name=arquivo)

        else:
            resultado = comparar_ia(t1, t2)
            st.text_area("Resultado IA", resultado, height=400)

    else:
        st.warning("Envie os dois PDFs.")