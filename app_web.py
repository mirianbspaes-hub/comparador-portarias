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
    """
    Transforma o Markdown da IA em formatação profissional do Word.
    ~~texto~~ -> Riscado e Vermelho (Texto antigo/excluído)
    **texto** -> Negrito e Azul Escuro (Texto novo/incluído)
    """
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    for linha in texto_ia.split('\n'):
        linha = linha.strip()
        if not linha:
            doc.add_paragraph() # Espaço entre parágrafos
            continue
            
        p = doc.add_paragraph()
        
        # Regex para identificar os marcadores de alteração da IA
        partes = re.split(r'(~~.*?~~|\*\*.*?\*\*)', linha)
        
        for parte in partes:
            if parte.startswith('~~') and parte.endswith('~~'):
                # FORMATO: REMOVIDO/ALTERADO (Riscado + Vermelho)
                texto_limpo = parte.replace('~~', '')
                run = p.add_run(texto_limpo)
                run.font.strike = True
                run.font.color.rgb = RGBColor(200, 0, 0)
            elif parte.startswith('**') and parte.endswith('**'):
                # FORMATO: NOVO (Negrito + Azul)
                texto_limpo = parte.replace('**', '')
                run = p.add_run(texto_limpo)
                run.bold = True
                run.font.color.rgb = RGBColor(0, 51, 102)
            else:
                # TEXTO QUE PERMANECE IGUAL
                p.add_run(parte)
                
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def processar_comparacao_ia(texto_base, texto_alteracoes):
    """
    Instrução crucial: O Texto 1 é a BASE INTEGRAL. O Texto 2 são apenas os comandos de mudança.
    """
    
    prompt_sistema = """
    Você é um especialista em Consolidação Normativa Jurídica.
    
    SUA TAREFA:
    Você deve pegar o 'TEXTO 1 (BASE INTEGRAL)' e usá-lo como o corpo principal do documento. 
    Você percorrerá o TEXTO 1 e, somente onde o 'TEXTO 2 (ALTERAÇÕES)' indicar uma mudança, você aplicará a alteração no local exato.
    
    REGRAS DE FORMATAÇÃO:
    1. NÃO SUPRIMA NADA do Texto 1 que não tenha sido expressamente alterado. O resultado final deve ser a Portaria completa.
    2. Onde houver ALTERAÇÃO: coloque o texto original do Texto 1 riscado como ~~texto antigo~~ e, logo abaixo, a nova redação em negrito como **texto novo**.
    3. Onde houver INCLUSÃO: insira o novo parágrafo/artigo no local correto em negrito **texto novo**.
    4. Ao final de cada alteração, adicione a nota de rodapé jurídica (ex: Redação dada pela Portaria nº X).
    5. Mantenha a estrutura original de Artigos, Parágrafos, Incisos e Alíneas.
    """

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": prompt_sistema},
                {"role": "user", "content": f"TEXTO 1 (BASE INTEGRAL QUE DEVE SER MANTIDA):\n{texto_base[:15000]}\n\nTEXTO 2 (SOMENTE AS ALTERAÇÕES A APLICAR):\n{texto_alteracoes[:10000]}"}
            ],
            temperature=0
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Erro na comunicação com a IA: {str(e)}"

# =========================
# INTERFACE STREAMLIT
# =========================

st.set_page_config(page_title="Comparador SAT - MTE", layout="wide")
st.title("⚖️ Consolidador de Portarias (Base Integral)")

st.warning("⚠️ O sistema usará o primeiro PDF como base completa e aplicará as mudanças contidas no segundo PDF.")

col1, col2 = st.columns(2)
with col1:
    pdf_base = st.file_uploader("1. Carregar Portaria ANTIGA COMPLETA (Base)", type="pdf")
with col2:
    pdf_alt = st.file_uploader("2. Carregar Documento de ALTERAÇÕES (Texto Novo)", type="pdf")

if st.button("🚀 Gerar Portaria Consolidada"):
    if not pdf_base or not pdf_alt:
        st.error("Upload obrigatório dos dois arquivos.")
    elif not client:
        st.error("API Key não configurada.")
    else:
        with st.spinner("Consolidando textos..."):
            t_base = extrair_texto_pdf(pdf_base)
            t_alt = extrair_texto_pdf(pdf_alt)
            
            # Chama a função que trata o Texto 1 como soberano/base
            resultado_ia = processar_comparacao_ia(t_base, t_alt)
            
            if "Erro" in resultado_ia:
                st.error(resultado_ia)
            else:
                arquivo_docx = gerar_word_com_estilo(resultado_ia)
                st.success("✅ Portaria consolidada com sucesso!")
                
                st.download_button(
                    label="📥 Baixar Portaria_Consolidada.docx",
                    data=arquivo_docx,
                    file_name="Portaria_Consolidada.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                with st.expander("Prévia das Alterações"):
                    st.write(resultado_ia)