import streamlit as st
import pdfplumber
import re
import io
import os
import time
from docx import Document
from docx.shared import Pt, RGBColor
from openai import OpenAI

# Configuração da API
api_key = st.secrets.get("OPENAI_API_KEY") if "OPENAI_API_KEY" in st.secrets else os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=api_key) if api_key else None

# ==========================================
# 1. LIMPEZA E EXTRAÇÃO DE TEXTO
# ==========================================
def limpar_texto_pdf(texto):
    """Remove cabeçalhos, rodapés e links chatos do PDF do DOU."""
    # Remove URLs
    texto = re.sub(r'https?://\S+', '', texto)
    # Remove padrões de data/hora e numeração de página (ex: 31/03/2026, 1/107)
    texto = re.sub(r'\d{2}/\d{2}/\d{4}, \d{2}:\d{2}', '', texto)
    texto = re.sub(r'\d+/\d{3,}', '', texto)
    # Remove termos fixos de cabeçalho do sistema
    texto = re.sub(r'Portaria - Portaria - DOU - Imprensa Nacional', '', texto, flags=re.IGNORECASE)
    # Remove excesso de quebras de linha
    texto = re.sub(r'\n\s*\n', '\n\n', texto)
    return texto.strip()

def extrair_texto_pdf(pdf):
    texto = ""
    with pdfplumber.open(pdf) as p:
        for page in p.pages:
            texto += (page.extract_text() or "") + "\n"
    return limpar_texto_pdf(texto)

# ==========================================
# 2. LÓGICA DE CONSOLIDAÇÃO
# ==========================================
def processar_consolidacao(texto_base, texto_alteracoes):
    # Divide o texto base em blocos menores (aprox. 3-4 artigos por vez)
    blocos = []
    partes = re.split(r'(\nArt\. \d+)', texto_base)
    
    # Reagrupa os artigos para não quebrar no meio de um parágrafo
    temp_bloco = ""
    for p in partes:
        if len(temp_bloco) + len(p) < 6000:
            temp_bloco += p
        else:
            blocos.append(temp_bloco)
            temp_bloco = p
    blocos.append(temp_bloco)

    resultado_final = []
    barra_progresso = st.progress(0)
    status = st.empty()

    prompt_sistema = """Você é um Consultor Jurídico sênior. Sua tarefa é CONSOLIDAR normas.
    REGRAS CRÍTICAS:
    1. Se o 'TEXTO DE ALTERAÇÕES' não mencionar nenhum artigo presente no 'BLOCO ORIGINAL', responda EXATAMENTE: MANTER_ORIGINAL.
    2. Se houver alteração: SUBSTITUA o texto original. 
    3. Use riscado para o que saiu: ~~texto antigo~~ (Revogado pela [[Portaria X]]).
    4. Use negrito para o que entrou: **texto novo** (Alterado pela [[Portaria X]]).
    5. NÃO duplique artigos. Se o Art. 6º mudou, mostre apenas UMA VEZ o Art. 6º com as marcações de alteração dentro dele.
    6. Ignore cabeçalhos, rodapés ou links que tenham sobrado no texto."""

    for i, bloco in enumerate(blocos):
        if not bloco.strip(): continue
        
        status.markdown(f"**Analisando bloco {i+1} de {len(blocos)}...**")
        
        try:
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": prompt_sistema},
                    {"role": "user", "content": f"BLOCO ORIGINAL:\n{bloco}\n\nTEXTO DE ALTERAÇÕES:\n{texto_alteracoes}"}
                ],
                temperature=0
            )
            resposta = response.choices[0].message.content.strip()
            
            if "MANTER_ORIGINAL" in resposta:
                resultado_final.append(bloco)
            else:
                resultado_final.append(resposta)
        except Exception as e:
            resultado_final.append(bloco)
            st.error(f"Erro no bloco {i}: {e}")
        
        barra_progresso.progress((i + 1) / len(blocos))
    
    status.empty()
    return "\n\n".join(resultado_final)

# ==========================================
# 3. GERADOR DE WORD (ESTILIZADO)
# ==========================================
def gerar_word(texto_final):
    doc = Document()
    # Configuração de Fonte Padrão
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    for linha in texto_final.split('\n'):
        if not linha.strip():
            doc.add_paragraph()
            continue
            
        p = doc.add_paragraph()
        # Regex para identificar as marcações da IA
        partes = re.split(r'(~~.*?~~|\*\*.*?\*\*|\[\[.*?\]\])', linha)
        
        for parte in partes:
            if parte.startswith('~~') and parte.endswith('~~'):
                run = p.add_run(parte[2:-2])
                run.font.strike = True
                run.font.color.rgb = RGBColor(150, 150, 150) # Cinza para o revogado
            elif parte.startswith('**') and parte.endswith('**'):
                run = p.add_run(parte[2:-2])
                run.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0)
            elif parte.startswith('[[') and parte.endswith(']]'):
                run = p.add_run(parte[2:-2])
                run.font.color.rgb = RGBColor(0, 0, 255) # Azul para referência de portaria
                run.italic = True
            else:
                p.add_run(parte)
                
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# 4. INTERFACE
# ==========================================
st.set_page_config(page_title="Consolidador Pro", page_icon="⚖️")
st.title("⚖️ Consolidador Jurídico Inteligente")
st.markdown("Este app remove sujeiras de PDF e garante que artigos não sejam duplicados.")

pdf_orig = st.file_uploader("Suba a Portaria ORIGINAL", type="pdf")
pdf_alt = st.file_uploader("Suba a Portaria que ALTERA", type="pdf")

if st.button("🚀 Consolidar Agora"):
    if pdf_orig and pdf_alt:
        with st.spinner("Limpando PDFs e comparando textos..."):
            t_orig = extrair_texto_pdf(pdf_orig)
            t_alt = extrair_texto_pdf(pdf_alt)
            
            # Processamento
            texto_consolidado = processar_consolidacao(t_orig, t_alt)
            arquivo_word = gerar_word(texto_consolidado)
            
            st.success("Pronto! Consolidação finalizada.")
            st.download_button(
                label="📥 Baixar Word Consolidado",
                data=arquivo_word,
                file_name="Portaria_Consolidada_Final.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            with st.expander("Ver prévia do texto"):
                st.write(texto_consolidado)
    else:
        st.warning("Por favor, envie os dois arquivos.")