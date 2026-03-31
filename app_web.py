import streamlit as st
import pdfplumber
import re
import io
import os
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
        
        # Identifica: Riscados (~~), Negritos (**) e Links ([[ ]])
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
    # AJUSTE CRÍTICO: Instruções anti-alucinação e fidelidade aos dados
    prompt_sistema = """
    Você é um compilador jurídico de alta precisão. Sua tarefa é integrar alterações em uma norma base.

    REGRAS DE FIDELIDADE (MUITO IMPORTANTE):
    1. NÃO adicione leis, datas ou nomes que NÃO estejam nos textos fornecidos. 
    2. NÃO invente fundamentações legais (ex: Lei nº 14.967). Se não está no texto, não coloque.
    3. Use o 'TEXTO 1' como base integral. O 'TEXTO 2' contém as únicas alterações permitidas.

    REGRAS DE FORMATAÇÃO:
    - Texto removido do TEXTO 1: Use ~~texto antigo~~ (Revogado pela [[Nome da Portaria do Texto 2]]).
    - Texto novo vindo do TEXTO 2: Use **novo texto** (Incluído pela [[Nome da Portaria do Texto 2]]).
    - Nome da Portaria: Sempre em colchetes duplos [[Portaria nº XXX]].
    - NÃO use notas de rodapé. Coloque a citação imediatamente após a alteração.
    """

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": prompt_sistema},
                {"role": "user", "content": f"TEXTO 1 (BASE):\n{texto_base[:15000]}\n\nTEXTO 2 (ALTERAÇÕES):\n{texto_alteracoes[:10000]}"}
            ],
            temperature=0 # Temperatura 0 evita que a IA invente coisas (alucinação)
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Erro na IA: {str(e)}"

# =========================
# INTERFACE STREAMLIT
# =========================

st.set_page_config(page_title="Consolidador SAT", layout="wide")
st.title("⚖️ Consolidador de Portarias Profissional")

col1, col2 = st.columns(2)
with col1:
    pdf_base = st.file_uploader("1. Portaria ANTIGA (Base Integral)", type="pdf")
with col2:
    pdf_alt = st.file_uploader("2. Documento de ALTERAÇÕES", type="pdf")

if st.button("🚀 Gerar Portaria Consolidada"):
    if pdf_base and pdf_alt:
        with st.spinner("Consolidando..."):
            t_base = extrair_texto_pdf(pdf_base)
            t_alt = extrair_texto_pdf(pdf_alt)
            
            resultado_ia = processar_comparacao_ia(t_base, t_alt)
            
            if "Erro" not in resultado_ia:
                arquivo_docx = gerar_word_com_estilo(resultado_ia)
                st.success("✅ Documento consolidado!")
                
                st.download_button(
                    label="📥 Baixar Portaria_Consolidada.docx",
                    data=arquivo_docx,
                    file_name="Portaria_Consolidada.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.error(resultado_ia)