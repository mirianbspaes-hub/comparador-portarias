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

def dividir_texto(texto, max_chars=8000):
    """Divide textos grandes em blocos menores para a IA não se perder nem resumir."""
    blocos = []
    inicio = 0
    while inicio < len(texto):
        fim = inicio + max_chars
        if fim < len(texto):
            # Tenta quebrar em dupla quebra de linha ou quebra simples
            pos_quebra = texto.rfind('\n\n', inicio, fim)
            if pos_quebra == -1 or pos_quebra <= inicio:
                pos_quebra = texto.rfind('\n', inicio, fim)
            if pos_quebra != -1 and pos_quebra > inicio:
                fim = pos_quebra
        blocos.append(texto[inicio:fim])
        inicio = fim
    return blocos

def gerar_word_fidelidade_total(texto_ia):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)
    style.paragraph_format.space_after = Pt(6)
    style.paragraph_format.line_spacing = 1.0 

    for linha in texto_ia.split('\n'):
        linha = linha.strip()
        
        if not linha:
            doc.add_paragraph()
            continue
            
        p = doc.add_paragraph()
        
        # Preserva recuos legais de artigos e incisos
        if linha.startswith("Art."):
            p.paragraph_format.first_line_indent = Pt(0)
        elif linha.startswith("§") or re.match(r'^[a-z]\)|\d+\.|[IVXLC]+\s?-', linha):
            p.paragraph_format.left_indent = Pt(36)

        partes = re.split(r'(~~.*?~~|\*\*.*?\*\*|\[\[.*?\]\])', linha)
        
        for parte in partes:
            if parte.startswith('~~') and parte.endswith('~~'):
                # TEXTO REMOVIDO: Riscado e Preto
                texto_limpo = parte.replace('~~', '')
                run = p.add_run(texto_limpo)
                run.font.strike = True
                run.font.color.rgb = RGBColor(0, 0, 0)
            elif parte.startswith('**') and parte.endswith('**'):
                # TEXTO NOVO: Negrito e Preto
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
                run = p.add_run(parte)
                run.font.color.rgb = RGBColor(0, 0, 0)
                
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def processar_comparacao_ia_blocos(texto_base, texto_alteracoes):
    blocos = dividir_texto(texto_base)
    resultado_final = ""
    
    total_blocos = len(blocos)
    barra_progresso = st.progress(0)
    status_texto = st.empty()

    prompt_sistema = """
    Você é um compilador jurídico de precisão absoluta. 
    Sua única tarefa é aplicar as alterações do 'TEXTO 2' no 'BLOCO DO TEXTO 1'.

    REGRAS DE FIDELIDADE (PROIBIÇÕES CRÍTICAS):
    1. NÃO MESCLE PREÂMBULOS: Ignore o preâmbulo ou cabeçalho do TEXTO 2. Aplique APENAS os comandos de alteração (ex: "O art. X passa a vigorar...").
    2. NÃO ADICIONE NADA que não seja uma ordem direta do TEXTO 2.
    3. SE O TEXTO 2 NÃO MANDAR ALTERAR NADA NESTE BLOCO ESPECÍFICO, DEVOLVA O BLOCO EXATAMENTE COMO ELE É. Não resuma.
    
    REGRAS DE FORMATAÇÃO:
    - Onde houver alteração, mantenha o original riscado: ~~texto antigo~~ (Revogado pela [[Nome da Portaria do Texto 2]]).
    - Insira a nova redação logo abaixo: **texto novo** (Incluído pela [[Nome da Portaria do Texto 2]]).
    - Use os colchetes duplos [[ ]] estritamente para o nome da portaria alteradora.
    """

    for i, bloco in enumerate(blocos):
        status_texto.markdown(f"**Analisando parte {i+1} de {total_blocos} da Portaria...**")
        
        try:
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": prompt_sistema},
                    {"role": "user", "content": f"BLOCO DO TEXTO 1 (BASE):\n{bloco}\n\nTEXTO 2 (ALTERAÇÕES A APLICAR SE NECESSÁRIO):\n{texto_alteracoes[:15000]}"}
                ],
                temperature=0
            )
            resultado_final += response.choices[0].message.content + "\n\n"
        except Exception as e:
            return f"Erro na comunicação com a IA na parte {i+1}: {str(e)}"
            
        barra_progresso.progress((i + 1) / total_blocos)
        time.sleep(1) # Pausa curta para não estourar o limite da API (Rate Limit)

    status_texto.empty()
    return resultado_final

# =========================
# INTERFACE STREAMLIT
# =========================

st.set_page_config(page_title="Consolidador Profissional", layout="wide")
st.title("⚖️ Consolidador de Portarias Longas e Complexas")

# Estado da sessão para evitar erros de renderização (removeChild)
if 'resultado_consolidado' not in st.session_state:
    st.session_state.resultado_consolidado = None

col1, col2 = st.columns(2)
with col1:
    pdf_base = st.file_uploader("1. Portaria ORIGINAL (Base Completa)", type="pdf", key="f_base")
with col2:
    pdf_alt = st.file_uploader("2. Portaria ALTERADORA", type="pdf", key="f_alt")

if st.button("🚀 Processar e Consolidar (Lê Anexos e Tabelas)", key="btn_run"):
    if not pdf_base or not pdf_alt:
        st.warning("Faça o upload dos dois arquivos PDF.")
    elif not client:
        st.error("API Key da OpenAI não configurada.")
    else:
        t_base = extrair_texto_pdf(pdf_base)
        t_alt = extrair_texto_pdf(pdf_alt)
        
        # Limpa o resultado anterior da tela se houver
        st.session_state.resultado_consolidado = None
        
        with st.spinner("Iniciando a leitura estruturada do documento..."):
            st.session_state.resultado_consolidado = processar_comparacao_ia_blocos(t_base, t_alt)

# Container isolado para o resultado
if st.session_state.resultado_consolidado:
    res = st.session_state.resultado_consolidado
    if "Erro" not in res:
        st.success("✅ Consolidação de todas as páginas concluída com sucesso!")
        doc_buffer = gerar_word_fidelidade_total(res)
        
        st.download_button(
            label="📥 Baixar Portaria_Consolidada.docx",
            data=doc_buffer,
            file_name="Portaria_Consolidada_Final.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="btn_dl_final"
        )
        
        with st.expander("Visualizar texto consolidado pelo sistema"):
            st.write(res)
    else:
        st.error(res)