import streamlit as st
import PyPDF2
import re
from io import BytesIO
from openpyxl import load_workbook

def extrair_dados_pdf(pdf_file):
    try:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        texto = ""
        for page in pdf_reader.pages:
            texto += page.extract_text() + "\n"
        return texto
    except Exception as e:
        st.error(f"Erro ao ler o PDF: {str(e)}")
        return ""

def identificar_modelo(texto):
    texto = texto.upper()
    if "FILME" in texto:
        return "filme"
    elif "SACO" in texto or "SACOS" in texto:
        return "saco"
    else:
        return "saco"  # padrão caso não identifique

def preencher_celula(ws, texto_busca, valor, col_offset=1):
    """Preenche a célula ao lado do texto encontrado"""
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and str(texto_busca).strip() in str(cell.value).strip():
                ws.cell(row=cell.row, column=cell.column+col_offset, value=valor)
                return True
    return False

def preencher_planilha_seguro(modelo, texto_extraido):
    try:
        # Carrega a planilha correta
        planilha_path = "SACO.xlsx" if modelo == "saco" else "FILME.xlsx"
        wb = load_workbook(planilha_path)
        ws = wb.active

        # Mostra o texto extraído para debug (opcional)
        with st.expander("Visualizar texto extraído do PDF"):
            st.text(texto_extraido[:2000] + ("..." if len(texto_extraido) > 2000 else ""))

        # Dicionário de mapeamento de campos
        campos = {
            "cliente": r"CLIENTE[:\s]*(.*)",
            "produto": r"PRODUTO[:\s]*(.*)",
            "codigo": r"(?:CÓDIGO|CÓD\.|COD)[\s:]*(.*)",
            "largura": r"LARGURA[ (A-Z)]*[:\s]*(\d+[,.]?\d*)",
            "comprimento": r"COMPRIMENTO[ (A-Z)]*[:\s]*(\d+[,.]?\d*)", 
            "espessura": r"ESPESSURA[ (A-Z)]*[:\s]*(\d+[,.]?\d*)",
            "passo": r"PASSO[ (A-Z)]*[:\s]*(\d+[,.]?\d*)",
            "observacoes": r"OBSERVAÇÕES[\s\n]*(.*?)(?=\n\s*\n|$)"
        }

        # Extrai os dados
        dados = {}
        for campo, regex in campos.items():
            match = re.search(regex, texto_extraido, re.IGNORECASE)
            if match:
                dados[campo] = match.group(1).strip()

        # Mostra os dados encontrados
        st.write("Dados identificados:")
        st.json(dados)

        # Preenche a planilha
        if modelo == "saco":
            preencher_celula(ws, "1.1 CLIENTE:", dados.get("cliente", ""))
            preencher_celula(ws, "1.3 PRODUTO:", dados.get("produto", ""))
            preencher_celula(ws, "1.2 CÓD. PRODUTO:", dados.get("codigo", ""))
            preencher_celula(ws, "2.3 LARGURA", dados.get("largura", ""))
            preencher_celula(ws, "2.4 COMPRIMENTO", dados.get("comprimento", ""))
            preencher_celula(ws, "2.5 ESPESSURA", dados.get("espessura", ""))
        else:  # filme
            preencher_celula(ws, "1.1 CLIENTE:", dados.get("cliente", ""))
            preencher_celula(ws, "1.3 PRODUTO:", dados.get("produto", ""))
            preencher_celula(ws, "1.2 CÓD. PRODUTO:", dados.get("codigo", ""))
            preencher_celula(ws, "2.3 LARGURA", dados.get("largura", ""))
            preencher_celula(ws, "2.4 PASSO DA FOTOCÉLULA", dados.get("passo", ""))
            preencher_celula(ws, "2.5 ESPESSURA", dados.get("espessura", ""))
        
        # Sempre tenta preencher observações
        preencher_celula(ws, "OBSERVAÇÕES", dados.get("observacoes", ""))

        # Salva a planilha
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    except Exception as e:
        st.error(f"Erro ao processar a planilha: {str(e)}")
        return None

# Interface do usuário
st.title("📋 Gerador de Fichas Técnicas")
st.markdown("**SACOS e FILMES**")

uploaded_file = st.file_uploader("Envie o PDF da ficha", type=["pdf"])

if uploaded_file:
    texto = extrair_dados_pdf(uploaded_file)
    if texto:
        modelo = identificar_modelo(texto)
        st.success(f"Modelo detectado: {modelo.upper()}")
        
        planilha = preencher_planilha_seguro(modelo, texto)
        
        if planilha:
            st.download_button(
                label="⬇️ Baixar Planilha Preenchida",
                data=planilha,
                file_name=f"FICHA_{modelo.upper()}_PREENCHIDA.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Planilha gerada com sucesso!")
        else:
            st.error("Falha ao gerar a planilha. Verifique os dados extraídos.")
