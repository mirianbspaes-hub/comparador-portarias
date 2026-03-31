import streamlit as st
import pdfplumber
import re
import io
import os
import time
from docx import Document
from docx.shared import Pt, RGBColor
from openai import OpenAI

# =========================
# CONFIGURAÇÃO DE AMBIENTE
# =========================
api_key = st.secrets.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=api_key) if api_key else None

# =========================
# FUNÇÕES DE PROCESSAMENTO
# =========================

def extrair_texto_pdf(pdf):
    texto = ""
    with pdfplumber.open(pdf) as p:
        for page in p.pages:
            texto += (page.extract_text() or "") + "\n"
    return texto

def gerar_word_com_estilo(texto_ia):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    for linha in texto_ia.split('\n'):
        linha = linha.strip()
        if not linha:
            doc.add_paragraph()
            continue
            
        p = doc.add_paragraph()
        partes = re.split(r'(~~.*?~~|\*\*.*?\*\*|\[\[.*?\]\])', linha)
        
        for parte in partes:
            if parte.startswith('~~') and parte.endswith('~~'):
                texto_limpo = parte.replace('~~', '')
                run = p.add_run(texto_limpo)
                run.font.strike = True
            elif parte.startswith('**') and parte.endswith('**'):
                texto_limpo = parte.replace('**', '')
                run = p.add_run(texto_limpo)
                run.bold = True
            elif parte.startswith('[[') and parte.endswith(']]'):
                texto_limpo = parte.replace('[[', '').replace(']]', '')
                run = p.add_run(texto_limpo)
                run.font.color.rgb = RGBColor(0, 0, 255)
                run.underline = True
            else:
                p.add_run(parte)
                
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def processar_comparacao_ia(texto_base, texto_alteracoes):
    prompt_sistema = """
    Você é um compilador jurídico de precisão absoluta. 
    Sua tarefa é integrar as alterações do 'TEXTO 2' na norma base 'TEXTO 1'.

    REGRAS DE OURO:
    1. FIDELIDADE TOTAL: NÃO adicione leis, datas, artigos ou fatos que NÃO constem nos arquivos enviados. 
    2. BASE INTEGRAL: Use o 'TEXTO 1' como estrutura completa. Não suprima artigos não alterados.
    3. SEM NOTAS DE RODAPÉ: A citação da portaria deve vir logo após a alteração.
    
    FORMATAÇÃO:
    - Removido: ~~texto original~~ (Revogado pela [[Nome da Portaria do Texto 2]]).
    - Novo: **novo texto** (Incluído pela [[Nome da Portaria do Texto 2]]).
    - Portaria: Sempre entre colchetes duplos [[Portaria nº XXX]].
    """

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": prompt_sistema},
                {"role": "user", "content": f"TEXTO 1 (BASE):\n{texto_base[:18000]}\n\nTEXTO 2 (ALTERAÇÕES):\n{texto_alteracoes[:12000]}"}
            ],
            temperature=0
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Erro na IA: {str(e)}"

# =========================
# INTERFACE STREAMLIT
# =========================

st.set_page_config(page_title="Consolidador SAT", layout="wide")
st.title("⚖️ Consolidador de Portarias Profissional")

# Inicializa o estado para evitar erros de renderização
if 'resultado_ia' not in st.session_state:
    st.session_state.resultado_ia = None

col1, col2 = st.columns(2)
with col1:
    pdf_base = st.file_uploader("1. Portaria ANTIGA (Base)", type="pdf")
with col2:
    pdf_alt = st.file_uploader("2. Portaria de ALTERAÇÕES", type="pdf")

if st.button("🚀 Gerar Portaria Consolidada"):
    if pdf_base and pdf_alt:
        with st.spinner("Processando..."):
            t_base = extrair_texto_pdf(pdf_base)
            t_alt = extrair_texto_pdf(pdf_alt)
            st.session_state.resultado_ia = processar_comparacao_ia(t_base, t_alt)
    else:
        st.error("Carregue os dois arquivos.")

# Área de download protegida para evitar erro de 'removeChild'
if st.session_state.resultado_ia:
    with st.container():
        res = st.session_state.resultado_ia
        if "Erro" not in res:
            doc_buffer = gerar_word_com_estilo(res)
            st.success("✅ Comparação pronta!")
            
            st.download_button(
                label="📥 Baixar Portaria Consolidada",
                data=doc_buffer,
                file_name="Portaria_Consolidada.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="btn_download_v1" # Chave única para estabilidade
            )
            
            with st.expander("Ver prévia do texto"):
                st.write(res)
        else:
            st.error(res)