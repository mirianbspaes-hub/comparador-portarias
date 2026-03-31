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
    Regras Aplicadas:
    - ~~texto~~ -> Riscado (Cor Preta)
    - **texto** -> Negrito (Cor Preta)
    - [[Portaria XXX]] -> Azul e Sublinhado (Link)
    """
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
        
        # Regex para identificar: Riscados (~~), Negritos (**) e Links ([[ ]])
        partes = re.split(r'(~~.*?~~|\*\*.*?\*\*|\[\[.*?\]\])', linha)
        
        for parte in partes:
            if parte.startswith('~~') and parte.endswith('~~'):
                # TEXTO REMOVIDO: Riscado em preto
                texto_limpo = parte.replace('~~', '')
                run = p.add_run(texto_limpo)
                run.font.strike = True
                run.font.color.rgb = RGBColor(0, 0, 0)
            elif parte.startswith('**') and parte.endswith('**'):
                # TEXTO ACRESCENTADO: Negrito em preto
                texto_limpo = parte.replace('**', '')
                run = p.add_run(texto_limpo)
                run.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0)
            elif parte.startswith('[[') and parte.endswith(']]'):
                # NOME DA PORTARIA: Azul e Sublinhado
                texto_limpo = parte.replace('[[', '').replace(']]', '')
                run = p.add_run(texto_limpo)
                run.font.color.rgb = RGBColor(0, 0, 255)
                run.underline = True
            else:
                # TEXTO NORMAL (PRETO)
                run = p.add_run(parte)
                run.font.color.rgb = RGBColor(0, 0, 0)
                
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def processar_comparacao_ia(texto_base, texto_alteracoes):
    prompt_sistema = """
    Você é um especialista em Consolidação Normativa Jurídica.
    
    SUA TAREFA:
    Use o 'TEXTO 1' como base integral e aplique as mudanças do 'TEXTO 2'.
    
    REGRAS DE FORMATAÇÃO E CONTEÚDO:
    1. NÃO crie seções de "Nota de rodapé".
    2. Texto alterado/removido: Use ~~texto antigo~~ (Revogado pela [[Nome da Portaria]]).
    3. Texto acrescentado: Use **novo texto** (Incluído pela [[Nome da Portaria]]).
    4. A citação da portaria deve vir IMEDIATAMENTE após a alteração, entre parênteses, na mesma linha ou logo abaixo, mas nunca como uma nota de rodapé isolada no fim da página.
    5. Para o NOME DA PORTARIA, use colchetes duplos: [[Portaria nº XXX, de data]].
    6. O texto base que não sofreu alteração deve ser mantido integralmente em fonte normal.
    7. Não use cores (vermelho/azul) no texto, exceto para o que estiver dentro de [[ ]].
    """

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": prompt_sistema},
                {"role": "user", "content": f"TEXTO 1 (BASE INTEGRAL):\n{texto_base[:15000]}\n\nTEXTO 2 (ALTERAÇÕES):\n{texto_alteracoes[:10000]}"}
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

col1, col2 = st.columns(2)
with col1:
    pdf_base = st.file_uploader("1. Portaria ANTIGA (Base Integral)", type="pdf")
with col2:
    pdf_alt = st.file_uploader("2. Documento de ALTERAÇÕES", type="pdf")

if st.button("🚀 Gerar Portaria Consolidada"):
    if pdf_base and pdf_alt:
        with st.spinner("Consolidando e formatando documento..."):
            t_base = extrair_texto_pdf(pdf_base)
            t_alt = extrair_texto_pdf(pdf_alt)
            
            resultado_ia = processar_comparacao_ia(t_base, t_alt)
            
            if "Erro" not in resultado_ia:
                arquivo_docx = gerar_word_com_estilo(resultado_ia)
                st.success("✅ Documento consolidado com sucesso!")
                
                st.download_button(
                    label="📥 Baixar Portaria_Consolidada.docx",
                    data=arquivo_docx,
                    file_name="Portaria_Consolidada.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                with st.expander("Prévia do Texto Gerado"):
                    st.write(resultado_ia)
            else:
                st.error(resultado_ia)
    else:
        st.error("Por favor, carregue os dois arquivos PDF.")