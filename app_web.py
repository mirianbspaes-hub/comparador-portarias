import streamlit as st
import pdfplumber
import re
import io
from docx import Document
from docx.shared import Pt, RGBColor
from openai import OpenAI

# =========================
# CONFIG
# =========================
api_key = st.secrets.get("OPENAI_API_KEY")
client = OpenAI(api_key=api_key) if api_key else None

# =========================
# EXTRAÇÃO
# =========================
def extrair_texto(pdf):
    texto = ""
    with pdfplumber.open(pdf) as p:
        for page in p.pages:
            texto += (page.extract_text() or "") + "\n"
    return texto

# =========================
# LIMPEZA
# =========================
def limpar_texto(texto):
    texto = re.sub(r"\.{5,}", "", texto)
    texto = re.sub(r"(\.\s*){5,}", "", texto)
    return texto

# =========================
# ESTRUTURAÇÃO
# =========================
def extrair_estrutura(texto):
    padrao = r"(Art\. ?\d+º?.*?)(?=Art\.|\Z)"
    artigos = re.findall(padrao, texto, re.DOTALL)

    estrutura = {}
    for art in artigos:
        num = re.search(r"Art\. ?(\d+)", art)
        if num:
            estrutura[f"Art. {num.group(1)}"] = art.strip()

    return estrutura

# =========================
# DETECTOR DE ALTERAÇÃO
# =========================
def detectar_tipo(alteracao):
    a = alteracao.lower()

    if "passa a vigorar" in a:
        return "substituicao"

    if "acrescenta" in a or "inclui" in a:
        return "inclusao"

    if "revoga" in a:
        return "revogacao"

    if "onde se lê" in a:
        return "parcial"

    return "desconhecido"

# =========================
# EXTRAIR ALTERAÇÕES
# =========================
def extrair_alteracoes(texto):
    blocos = re.split(r"Art\. ?\d+º?", texto)
    alteracoes = []

    for b in blocos:
        if len(b.strip()) > 50:
            tipo = detectar_tipo(b)
            art = re.search(r"art\. ?(\d+)", b.lower())

            if art:
                alteracoes.append({
                    "artigo": f"Art. {art.group(1)}",
                    "tipo": tipo,
                    "texto": b.strip()
                })

    return alteracoes

# =========================
# IA PARA CASOS COMPLEXOS
# =========================
def aplicar_ia(original, instrucao):
    if not client:
        return original

    prompt = """
Você é especialista em legislação brasileira.

Aplique a instrução ao texto original.

REGRAS:
- NÃO inventar
- NÃO resumir
- NÃO alterar estrutura
- manter fidelidade jurídica

FORMATAÇÃO:
- removido: ~~texto~~
- novo: **texto**
"""

    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": prompt},
            {"role": "user", "content": f"TEXTO:\n{original}\n\nINSTRUÇÃO:\n{instrucao}"}
        ],
        temperature=0
    )

    return resp.choices[0].message.content

# =========================
# CONSOLIDAÇÃO PRINCIPAL
# =========================
def consolidar(base, alteracoes):

    resultado = base.copy()

    for alt in alteracoes:
        art = alt["artigo"]
        tipo = alt["tipo"]
        texto = alt["texto"]

        if art not in resultado:
            continue

        original = resultado[art]

        # 🔥 SUBSTITUIÇÃO COMPLETA (SEM IA)
        if tipo == "substituicao":
            novo = re.split(r"redação:(.*)", texto, flags=re.DOTALL)
            if len(novo) > 1:
                resultado[art] = f"{art} ~~{original}~~\n{art} **{novo[1].strip()}**"
            else:
                resultado[art] = aplicar_ia(original, texto)

        # 🔥 INCLUSÃO
        elif tipo == "inclusao":
            resultado[art] += f"\n\n**{texto}**"

        # 🔥 REVOGAÇÃO
        elif tipo == "revogacao":
            resultado[art] = f"~~{original}~~"

        # 🔥 PARCIAL
        else:
            resultado[art] = aplicar_ia(original, texto)

    return resultado

# =========================
# GERAR WORD
# =========================
def gerar_word(textos):
    doc = Document()

    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    for art in textos.values():
        p = doc.add_paragraph()

        partes = re.split(r'(~~.*?~~|\*\*.*?\*\*)', art)

        for parte in partes:
            if parte.startswith("~~"):
                r = p.add_run(parte.replace("~~", ""))
                r.font.strike = True
                r.font.color.rgb = RGBColor(255, 0, 0)
            elif parte.startswith("**"):
                r = p.add_run(parte.replace("**", ""))
                r.bold = True
                r.font.color.rgb = RGBColor(0, 128, 0)
            else:
                p.add_run(parte)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# =========================
# UI
# =========================
st.set_page_config(page_title="Motor Jurídico Avançado", layout="wide")

st.title("⚖️ Consolidador Jurídico Profissional")

pdf1 = st.file_uploader("Portaria ORIGINAL", type="pdf")
pdf2 = st.file_uploader("Portaria ALTERADORA", type="pdf")

if st.button("🚀 Consolidar Documento"):

    if pdf1 and pdf2:

        base_texto = limpar_texto(extrair_texto(pdf1))
        alt_texto = limpar_texto(extrair_texto(pdf2))

        base = extrair_estrutura(base_texto)
        alteracoes = extrair_alteracoes(alt_texto)

        resultado = consolidar(base, alteracoes)

        doc = gerar_word(resultado)

        st.success("✅ Consolidação concluída!")

        st.download_button(
            "📥 Baixar Word",
            doc,
            "Portaria_Consolidada_Final.docx"
        )

    else:
        st.warning("Envie os dois PDFs.")