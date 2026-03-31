import streamlit as st
import pdfplumber
import re
from docx import Document
from docx.shared import Pt
import os
st.write("API KEY carregada:", bool(os.getenv("OPENAI_API_KEY")))
from openai import OpenAI

api_key = os.getenv("OPENAI_API_KEY")

if api_key:
    client = OpenAI(api_key=api_key)
    IA_ATIVA = True
else:
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


import re

def extrair_blocos_juridicos(texto):
    padrao = r"(Art\. ?\d+º?|§ ?\d+º?|[IVXLC]+\s?-|[a-z]\)|[0-9]+\.[0-9\.]+)"
    
    partes = re.split(padrao, texto)
    blocos = {}

    for i in range(1, len(partes), 2):
        chave = partes[i].strip()
        conteudo = partes[i+1].strip() if i+1 < len(partes) else ""
        blocos[chave] = conteudo

    return blocos


def comparar_juridico(b1, b2):
    resultado = []
    chaves = sorted(set(b1.keys()).union(b2.keys()))

    for c in chaves:
        t1 = b1.get(c, "").strip()
        t2 = b2.get(c, "").strip()

        if t1 == t2:
            resultado.append(("igual", c, t1))

        elif t1 and not t2:
            resultado.append(("removido", c, t1))

        elif not t1 and t2:
            resultado.append(("adicionado", c, t2))

        else:
            resultado.append(("alterado", c, t1, t2))

    return resultado


def gerar_docx_oficial(resultado):
    doc = Document()

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    for item in resultado:
        tipo = item[0]

        if tipo == "igual":
            chave, texto = item[1], item[2]
            doc.add_paragraph(f"{chave} {texto}")

        elif tipo == "removido":
            chave, texto = item[1], item[2]
            p = doc.add_paragraph()
            r = p.add_run(f"{chave} {texto}")
            r.font.strike = True

        elif tipo == "adicionado":
            chave, texto = item[1], item[2]
            p = doc.add_paragraph()
            r = p.add_run(f"{chave} {texto}")
            r.bold = True

        elif tipo == "alterado":
            chave, antigo, novo = item[1], item[2], item[3]

            # Texto antigo riscado
            p1 = doc.add_paragraph()
            r1 = p1.add_run(f"{chave} {antigo}")
            r1.font.strike = True

            # Novo texto logo abaixo
            p2 = doc.add_paragraph()
            r2 = p2.add_run(f"{chave} {novo}")
            r2.bold = True

    caminho = "Portaria_Comparada_Oficial.docx"
    doc.save(caminho)
    return caminho


def comparar_com_ia_refinada(t1, t2):

    prompt = f"""
Você é especialista em legislação brasileira.

Compare os textos abaixo seguindo RIGOROSAMENTE:

- NÃO resumir
- NÃO reescrever juridicamente
- NÃO alterar conteúdo
- NÃO omitir trechos

Faça:

1. Manter estrutura (Art., §, incisos)
2. Texto antigo riscado (~~texto~~)
3. Texto novo abaixo
4. Inclusões destacadas
5. Exclusões riscadas

Objetivo:
Gerar versão consolidada estilo Diário Oficial.

TEXTO ORIGINAL:
{t1}

TEXTO ALTERADO:
{t2}
"""

     # 🔁 retry automático (evita RateLimit)
    for tentativa in range(3):
        try:
            resp = client.chat.completions.create(
                model="gpt-4o-mini",  # 🔥 mais estável e barato
                messages=[{"role": "user", "content": prompt}],
                temperature=0
            )

            return resp.choices[0].message.content

        except Exception as e:
            if "RateLimit" in str(e):
                time.sleep(2 * (tentativa + 1))
            else:
                return f"Erro: {e}"

    return "❌ Muitas requisições. Tente novamente em alguns segundos."

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
            b1 = extrair_blocos_juridicos(t1)
            b2 = extrair_blocos_juridicos(t2)
            resultado = comparar_juridico(b1, b2)
            arquivo = gerar_docx_oficial(resultado)

            with open(arquivo, "rb") as f:
                st.download_button("📄 Baixar Word", f, file_name=arquivo)

        else:
            resultado = comparar_com_ia_refinada(t1, t2)
            st.text_area("Resultado IA", resultado, height=400)

    else:
        st.warning("Envie os dois PDFs.")