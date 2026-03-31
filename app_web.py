import streamlit as st
import pdfplumber
import re
import time
import io
import os
from docx import Document
from docx.shared import Pt, RGBColor
from openai import OpenAI

# =========================
# CONFIG IA
# =========================
api_key = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=api_key) if api_key else None

# =========================
# CONFIG STREAMLIT
# =========================
st.set_page_config(page_title="Comparador de Portarias", layout="wide")
st.title("⚖️ Comparador Profissional de Portarias")
st.write("Geração automática de comparativo jurídico com IA.")

# =========================
# FUNÇÕES DE TEXTO E WORD
# =========================

def extrair_texto(pdf):
    texto = ""
    with pdfplumber.open(pdf) as p:
        for page in p.pages:
            texto += (page.extract_text() or "") + "\n"
    return texto

def criar_docx_da_ia(texto_ia):
    """Converte o texto da IA (com markdown) para um documento Word formatado."""
    doc = Document()
    
    # Configuração de Estilo Padrão (Times New Roman 12)
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    linhas = texto_ia.split('\n')
    for linha in linhas:
        if not linha.strip():
            continue
            
        p = doc.add_paragraph()
        
        # Lógica para identificar texto riscado (~~texto~~) vindo da IA
        partes = re.split(r'(~~.*?~~)', linha)
        
        for parte in partes:
            if parte.startswith('~~') and parte.endswith('~~'):
                conteudo = parte.replace('~~', '')
                run = p.add_run(conteudo)
                run.font.strike = True
                run.font.color.rgb = RGBColor(200, 0, 0) # Vermelho para removido
            else:
                run = p.add_run(parte)
                # Se a linha parecer ser uma inclusão (pode-se ajustar o prompt da IA para marcar)
                if "(Redação dada por" in linha or "Incluído pela" in linha:
                    run.font.color.rgb = RGBColor(0, 50, 150) # Azul para referência legal
                
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def comparar_com_ia(t1, t2, tipo_doc):
    contexto = "Comparação integral de dois textos." if tipo_doc == "Comparação direta (2 textos completos)" else "Aplicação de novas redações sobre uma portaria original."
    
    prompt = f"""
    Você é um especialista em redação oficial e consolidação normativa.
    Tarefa: {contexto}
    
    REGRAS DE FORMATAÇÃO:
    1. Se um artigo/parágrafo foi alterado: Exiba o texto antigo totalmente riscado usando ~~texto antigo~~. 
    Logo abaixo, exiba o texto novo em sua forma normal.
    2. Se foi incluído: Exiba o texto novo e adicione ao final '(Incluído pela Portaria nº X)'.
    3. Se foi revogado: Exiba o texto antigo riscado ~~texto antigo~~ e adicione '(Revogado pela Portaria nº X)'.
    4. Mantenha a estrutura: Art., §, incisos, alíneas.
    
    TEXTO 1 (ORIGINAL):
    {t1}
    
    TEXTO 2 (ALTERADOR/NOVO):
    {t2}
    
    Retorne apenas o texto comparado, sem comentários extras.
    """

    resp = client.chat.completions.create(
        model="gpt-4o", # Recomendado para precisão jurídica
        messages=[{"role": "system", "content": "Você é um assistente jurídico de alta precisão."},
                  {"role": "user", "content": prompt}],
        temperature=0
    )
    return resp.choices[0].message.content

# =========================
# INTERFACE DO USUÁRIO
# =========================

col1, col2 = st.columns(2)
with col1:
    pdf_orig = st.file_uploader("📥 PDF ORIGINAL", type="pdf")
with col2:
    pdf_alt = st.file_uploader("📥 PDF ALTERADOR", type="pdf")

tipo_doc = st.radio("Tipo de documento:", [
    "Comparação direta (2 textos completos)", 
    "Portaria + Alterações"
])

if st.button("🚀 Processar e Gerar Word Automático"):
    if not pdf_orig or not pdf_alt:
        st.warning("Por favor, envie ambos os arquivos.")
    elif not client:
        st.error("Chave da API não encontrada nas variáveis de ambiente.")
    else:
        with st.spinner("Analisando documentos com IA..."):
            # 1. Extração
            texto_1 = extrair_texto(pdf_orig)
            texto_2 = extrair_texto(pdf_alt)
            
            # 2. Comparação via IA
            resultado_texto = comparar_com_ia(texto_1, texto_2, tipo_doc)
            
            # 3. Geração do Word
            arquivo_word = criar_docx_da_ia(resultado_texto)
            
            st.success("✅ Comparação concluída!")
            
            # 4. Área de Visualização e Download
            st.text_area("Prévia do Resultado:", resultado_texto, height=300)
            
            st.download_button(
                label="📄 Baixar Portaria Comparada (.docx)",
                data=arquivo_word,
                file_name="Portaria_Comparada_Atualizada.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )