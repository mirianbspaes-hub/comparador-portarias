import streamlit as st
import pdfplumber
import re
import io
from docx import Document
from docx.shared import Pt, RGBColor
from openai import OpenAI

# =========================
# CONFIGURAÇÃO
# =========================
api_key = st.secrets.get("OPENAI_API_KEY")
client = OpenAI(api_key=api_key) if api_key else None

def extrair_texto_pdf(pdf):
    texto = ""
    with pdfplumber.open(pdf) as p:
        for page in p.pages:
            texto += (page.extract_text() or "") + "\n"
    return texto

# NOVA FUNÇÃO: Divide o texto em blocos de aproximadamente 4000 palavras
def dividir_texto(texto, max_chars=12000):
    return [texto[i:i+max_chars] for i in range(0, len(texto), max_chars)]

def gerar_word_fidelidade(texto_ia):
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
                run = p.add_run(parte.replace('~~', ''))
                run.font.strike = True
            elif parte.startswith('**') and parte.endswith('**'):
                run = p.add_run(parte.replace('**', ''))
                run.bold = True
            elif parte.startswith('[[') and parte.endswith(']]'):
                run = p.add_run(parte.replace('[[', '').replace(']]', ''))
                run.font.color.rgb = RGBColor(0, 0, 255)
                run.underline = True
            else:
                p.add_run(parte)
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def comparar_por_blocos(t_base, t_alt):
    # Dividimos a portaria grande em pedaços
    blocos = dividir_texto(t_base)
    resultado_final = ""
    
    barra_progresso = st.progress(0)
    total_blocos = len(blocos)

    for i, bloco in enumerate(blocos):
        st.write(f"Processando parte {i+1} de {total_blocos}...")
        
        prompt = f"""
        Você é um compilador jurídico. 
        Mantenha INTEGRALMENTE o TEXTO BASE abaixo, mas aplique as ALTERAÇÕES se elas afetarem esta parte do texto.
        
        REGRAS:
        - Se algo mudou: ~~antigo~~ (Revogado pela [[Portaria]]) **novo** (Incluído pela [[Portaria]]).
        - Se não houver alteração nesta parte, apenas repita o texto original.
        - PROIBIDO inventar leis.
        """
        
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": prompt},
                {"role": "user", "content": f"TEXTO BASE (PARTE {i+1}):\n{bloco}\n\nALTERAÇÕES A VERIFICAR:\n{t_alt}"}
            ],
            temperature=0
        )
        resultado_final += response.choices[0].message.content + "\n"
        barra_progresso.progress((i + 1) / total_blocos)
        
    return resultado_final

# =========================
# INTERFACE
# =========================
st.set_page_config(page_title="Consolidador Portarias Grandes", layout="wide")
st.title("⚖️ Consolidador de Portarias Longas")

f1 = st.file_uploader("1. Portaria Base (Pode ser grande)", type="pdf")
f2 = st.file_uploader("2. Documento de Alterações", type="pdf")

if st.button("🚀 Iniciar Consolidação por Partes"):
    if f1 and f2:
        t1 = extrair_texto_pdf(f1)
        t2 = extrair_texto_pdf(f2)
        
        # Chama a função que processa bloco por bloco
        resultado = comparar_por_blocos(t1, t2)
        
        doc_buffer = gerar_word_fidelidade(resultado)
        st.success("✅ Concluído!")
        st.download_button("📥 Baixar Documento Completo", doc_buffer, "Portaria_Grande_Consolidada.docx")