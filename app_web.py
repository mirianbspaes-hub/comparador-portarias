import streamlit as st
import pdfplumber
import re
import io
import time
from docx import Document
from docx.shared import Pt, RGBColor
from openai import OpenAI

# =========================
# CONFIGURAÇÃO
# =========================
api_key = st.secrets.get("OPENAI_API_KEY")
client = OpenAI(api_key=api_key) if api_key else None

# =========================
# EXTRAÇÃO
# =========================
def limpar_reticencias(texto):
    # remove linhas compostas só por pontos
    texto = re.sub(r"\.{5,}", "", texto)

    # remove linhas com muitos pontos e espaços
    texto = re.sub(r"(\.\s*){5,}", "", texto)

    # remove linhas vazias geradas
    texto = "\n".join([l for l in texto.split("\n") if l.strip() != ""])

    return texto

def extrair_texto_pdf(pdf):
    texto = ""
    with pdfplumber.open(pdf) as p:
        for page in p.pages:
            texto += (page.extract_text() or "") + "\n"
    return texto

# =========================
# DIVISÃO EM BLOCOS
# =========================
def dividir_texto(texto, max_chars=12000):
    return [texto[i:i+max_chars] for i in range(0, len(texto), max_chars)]

# =========================
# GERAR WORD PROFISSIONAL
# =========================
def gerar_word_fidelidade(texto_ia):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    for linha in texto_ia.split('\n'):
        linha = linha.strip()

        if not linha:
            doc.add_paragraph()
            continue

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)

        partes = re.split(r'(~~.*?~~|\*\*.*?\*\*|\[\[.*?\]\])', linha)

        for parte in partes:
            if parte.startswith('~~') and parte.endswith('~~'):
                run = p.add_run(parte.replace('~~', ''))
                run.font.strike = True
                run.font.color.rgb = RGBColor(255, 0, 0)

            elif parte.startswith('**') and parte.endswith('**'):
                run = p.add_run(parte.replace('**', ''))
                run.bold = True
                run.font.color.rgb = RGBColor(0, 128, 0)

            elif parte.startswith('[[') and parte.endswith(']]'):
                run = p.add_run(parte.replace('[[', '').replace(']]', ''))
                run.font.color.rgb = RGBColor(0, 0, 255)
                run.underline = True

            else:
                p.add_run(parte)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# =========================
# IA COM CONTROLE
# =========================
def comparar_por_blocos(t_base, t_alt):

    if not client:
        st.error("❌ Configure a chave da OpenAI no Secrets.")
        return ""

    # limpar reticências antes de tudo
    t_base = limpar_reticencias(t_base)
    t_alt = limpar_reticencias(t_alt)

    blocos = dividir_texto(t_base)

    if len(blocos) > 20:
        st.warning("⚠️ Documento muito grande. Pode gerar custo alto.")

    resultado_final = ""
    barra_progresso = st.progress(0)
    total_blocos = len(blocos)

    for i, bloco in enumerate(blocos):

        st.write(f"Processando parte {i+1} de {total_blocos}...")

        # 🔥 FILTRO: só manda blocos relevantes
        if "Art." not in bloco and "§" not in bloco:
            resultado_final += bloco + "\n"
            barra_progresso.progress((i + 1) / total_blocos)
            continue

        prompt = """
Você é um especialista em consolidação normativa brasileira.

Sua tarefa:
Aplicar ALTERAÇÕES ao TEXTO BASE sem alterar conteúdo original desnecessariamente.

REGRAS OBRIGATÓRIAS:
- NÃO inventar conteúdo
- NÃO reescrever juridicamente
- NÃO resumir
- NÃO alterar estrutura (Art., §, incisos)
- Trechos com "....." indicam continuidade e NÃO são alteração

FORMATAÇÃO:
- Texto removido: ~~texto~~
- Texto novo: **texto**
- Indicar origem: [[Portaria]]

Se não houver alteração, repetir exatamente o texto original.
"""

        try:
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": prompt},
                    {"role": "user", "content": f"TEXTO BASE:\n{bloco}\n\nALTERAÇÕES:\n{t_alt}"}
                ],
                temperature=0
            )

            resultado_final += response.choices[0].message.content + "\n"

        except Exception as e:
            resultado_final += f"\n[ERRO NO BLOCO {i+1}: {e}]\n"

        barra_progresso.progress((i + 1) / total_blocos)
        time.sleep(0.5)

    return resultado_final

# =========================
# INTERFACE
# =========================
st.set_page_config(page_title="Consolidador de Portarias", layout="wide")

st.title("⚖️ Consolidador de Portarias (Profissional)")
st.write("Processa portarias grandes com IA, mantendo fidelidade jurídica.")

f1 = st.file_uploader("1. Portaria Base (completa)", type="pdf")
f2 = st.file_uploader("2. Documento de Alterações", type="pdf")

# =========================
# EXECUÇÃO
# =========================
if st.button("🚀 Iniciar Consolidação por Partes"):

    if f1 and f2:

        t1 = extrair_texto_pdf(f1)
        t2 = extrair_texto_pdf(f2)

        resultado = comparar_por_blocos(t1, t2)

        doc_buffer = gerar_word_fidelidade(resultado)

        st.success("✅ Consolidação concluída com sucesso!")

        st.download_button(
            "📥 Baixar Documento Consolidado",
            doc_buffer,
            "Portaria_Consolidada.docx"
        )

    else:
        st.warning("Envie os dois PDFs.")