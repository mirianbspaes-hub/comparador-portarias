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
    ~~texto~~ -> Riscado e Vermelho (Antigo)
    **texto** -> Negrito e Azul Escuro (Novo)
    """
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    for linha in texto_ia.split('\n'):
        linha = linha.strip()
        if not linha:
            # Adiciona espaço entre parágrafos, como no modelo
            doc.add_paragraph()
            continue
            
        p = doc.add_paragraph()
        
        # Regex para capturar os marcadores de alteração
        partes = re.split(r'(~~.*?~~|\*\*.*?\*\*)', linha)
        
        for parte in partes:
            if parte.startswith('~~') and parte.endswith('~~'):
                # FORMATO: REMOVIDO (Riscado + Vermelho)
                texto_limpo = parte.replace('~~', '')
                run = p.add_run(texto_limpo)
                run.font.strike = True
                run.font.color.rgb = RGBColor(200, 0, 0)
            elif parte.startswith('**') and parte.endswith('**'):
                # FORMATO: INCLUÍDO (Negrito + Azul)
                texto_limpo = parte.replace('**', '')
                run = p.add_run(texto_limpo)
                run.bold = True
                run.font.color.rgb = RGBColor(0, 51, 102)
            else:
                # FORMATO: TEXTO NORMAL/MANUTENÇÃO
                p.add_run(parte)
                
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def processar_comparacao_ia(t1, t2, tipo_doc):
    """
    Envia para a IA com instruções de consolidação normativa.
    """
    if tipo_doc == "Portaria + Alterações":
        contexto_adicional = "O Segundo Texto contém apenas as alterações que devem ser aplicadas sobre o Primeiro Texto."
    else:
        contexto_adicional = "Compare os dois textos completos e identifique as diferenças entre eles."

    prompt_sistema = f"""
    Você é um especialista em Consolidação Normativa Jurídica.
    Sua missão é criar um documento único que mostre a evolução do texto.
    
    REGRAS DE OURO:
    1. JAMAIS liste o texto original completo e depois o alterado.
    2. A comparação deve ser ITEM POR ITEM (Artigo, Parágrafo, Inciso).
    3. Se algo mudou: coloque a redação antiga usando ~~texto antigo~~ e, IMEDIATAMENTE ABAIXO, a nova redação usando **nova redação**.
    4. Se algo é novo: use **texto novo** e adicione (Incluído pela Portaria nº X).
    5. Se algo foi excluído: use ~~texto antigo~~ e adicione (Revogado pela Portaria nº X).
    6. Não resuma. Mantenha a literalidade jurídica.
    """

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": prompt_sistema},
                {"role": "user", "content": f"{contexto_adicional}\n\nTEXTO 1 (BASE):\n{t1[:15000]}\n\nTEXTO 2 (ALTERAÇÕES):\n{t2[:15000]}"}
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
st.title("⚖️ Sistema de Consolidação de Portarias")
st.info("Este app gera automaticamente um Word com textos riscados e novas redações destacadas.")

col1, col2 = st.columns(2)
with col1:
    pdf_base = st.file_uploader("1. Carregar Portaria Original (PDF)", type="pdf")
with col2:
    pdf_alt = st.file_uploader("2. Carregar Documento com Alterações (PDF)", type="pdf")

tipo_analise = st.radio(
    "Tipo de Documento:",
    ["Portaria + Alterações", "Comparação direta (2 textos completos)"],
    horizontal=True
)

if st.button("🚀 Gerar Documento Word Comparado"):
    if not pdf_base or not pdf_alt:
        st.warning("Por favor, faça o upload dos dois arquivos PDF para continuar.")
    elif not client:
        st.error("Erro: API Key da OpenAI não configurada.")
    else:
        with st.spinner("Analisando textos e aplicando regras jurídicas..."):
            # Passo 1: Extração
            texto_base = extrair_texto_pdf(pdf_base)
            texto_alt = extrair_texto_pdf(pdf_alt)
            
            # Passo 2: Inteligência de Comparação
            resultado_ia = processar_comparacao_ia(texto_base, texto_alt, tipo_analise)
            
            if "Erro" in resultado_ia:
                st.error(resultado_ia)
            else:
                # Passo 3: Geração do Word formatado
                arquivo_docx = gerar_word_com_estilo(resultado_ia)
                
                st.success("✅ Documento gerado com sucesso!")
                
                # Botão de Download
                st.download_button(
                    label="📥 Baixar Portaria_Comparada_Atualizada.docx",
                    data=arquivo_docx,
                    file_name="Portaria_Comparada_Atualizada.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                # Prévia Visual
                with st.expander("Visualizar alterações detectadas (Resumo)"):
                    st.markdown(resultado_ia.replace('~~', '~~').replace('**', '**'))