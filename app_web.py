import streamlit as st
import pdfplumber
import re
import time
import io
import os
from docx import Document
from docx.shared import Pt, RGBColor
from openai import OpenAI
import openai # Importação necessária para capturar o erro específico

# =========================
# CONFIG IA
# =========================
api_key = st.secrets.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=api_key) if api_key else None

# =========================
# FUNÇÕES DE TEXTO E WORD
# =========================

def extrair_texto(pdf):
    texto = ""
    try:
        with pdfplumber.open(pdf) as p:
            for page in p.pages:
                texto += (page.extract_text() or "") + "\n"
    except Exception as e:
        st.error(f"Erro ao ler PDF: {e}")
    return texto

def criar_docx_da_ia(texto_ia):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    linhas = texto_ia.split('\n')
    for linha in linhas:
        if not linha.strip():
            continue
        p = doc.add_paragraph()
        
        # Identifica o padrão de riscado da IA ~~texto~~
        partes = re.split(r'(~~.*?~~)', linha)
        for parte in partes:
            if parte.startswith('~~') and parte.endswith('~~'):
                conteudo = parte.replace('~~', '')
                run = p.add_run(conteudo)
                run.font.strike = True
                run.font.color.rgb = RGBColor(200, 0, 0)
            else:
                run = p.add_run(parte)
                if "Redação dada por" in linha or "Incluído pela" in linha:
                    run.font.color.rgb = RGBColor(0, 50, 150)
                
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def comparar_com_ia(t1, t2, tipo_doc):
    contexto = "Comparação integral" if tipo_doc == "Comparação direta (2 textos completos)" else "Aplicação de alterações"
    
    try:
        resp = client.chat.completions.create(
            model="gpt-4o-mini", # Alterado para evitar Rate Limit e reduzir custo
            messages=[
                {"role": "system", "content": "Você é um revisor jurídico. Compare os textos. Use ~~texto~~ para o que foi removido/alterado. Escreva a nova redação logo abaixo."},
                {"role": "user", "content": f"Tarefa: {contexto}\n\nORIGINAL:\n{t1[:15000]}\n\nALTERADO:\n{t2[:15000]}"} 
            ],
            temperature=0
        )
        return resp.choices[0].message.content
    except openai.RateLimitError:
        return "ERRO_LIMITE: Você atingiu o limite de velocidade da OpenAI. Aguarde um minuto ou verifique seus créditos no painel da OpenAI."
    except Exception as e:
        return f"ERRO_INESPERADO: {str(e)}"

# =========================
# INTERFACE
# =========================
st.set_page_config(page_title="Comparador de Portarias", layout="wide")
st.title("⚖️ Comparador Profissional de Portarias")

pdf_orig = st.file_uploader("📥 PDF ORIGINAL", type="pdf")
pdf_alt = st.file_uploader("📥 PDF ALTERADOR", type="pdf")

tipo_doc = st.radio("Tipo de documento:", [
    "Comparação direta (2 textos completos)", 
    "Portaria + Alterações"
])

if st.button("🚀 Processar e Gerar Word Automático"):
    if not pdf_orig or not pdf_alt:
        st.warning("Envie os dois arquivos.")
    elif not client:
        st.error("API Key não configurada.")
    else:
        with st.spinner("Analisando documentos..."):
            texto_1 = extrair_texto(pdf_orig)
            texto_2 = extrair_texto(pdf_alt)
            
            resultado_texto = comparar_com_ia(texto_1, texto_2, tipo_doc)
            
            if "ERRO_LIMITE" in resultado_texto:
                st.error(resultado_texto)
            elif "ERRO_INESPERADO" in resultado_texto:
                st.error(resultado_texto)
            else:
                arquivo_word = criar_docx_da_ia(resultado_texto)
                st.success("✅ Comparação concluída!")
                
                st.download_button(
                    label="📄 Baixar Portaria Comparada (.docx)",
                    data=arquivo_word,
                    file_name="Portaria_Comparada_Atualizada.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                st.text_area("Prévia:", resultado_texto, height=300)