import streamlit as st
import PyPDF2
import re
from io import BytesIO
from openpyxl import load_workbook
from pathlib import Path

EXCEL_PATH = Path(__file__).parent / "SACO.xlsx"

def extrair_dados(pdf_file):
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    return "\n".join(page.extract_text() for page in pdf_reader.pages)

def preencher_saco(texto):
    wb = load_workbook(EXCEL_PATH)
    ws = wb.active

    def extrair(padrao):
        resultado = re.search(padrao, texto, re.IGNORECASE)
        return resultado.group(1).strip() if resultado else ""

    ws['C7'] = extrair(r"NOME DO CLIENTE[:\s]*(.*)") or "COOPAVEL COOPERATIVA AGROINDUSTRIAL"
    ws['C8'] = extrair(r"PRODUTO[:\s]*(.*)") or "SACO MULTIUSO IMPRESSO TRANSPARENTE"
    ws['J7'] = extrair(r"PEDIDO N[º°:\s]*(.*)") or "36486"

    medidas = re.search(r"(\d+)X(\d+)X([\d,]+)", texto)
    if medidas:
        largura = float(medidas.group(1))
        altura = float(medidas.group(2))
        espessura = float(medidas.group(3).replace(",", "."))
        ws['A16'] = largura
        ws['E16'] = altura
        ws['I16'] = espessura
        ws['A20'] = largura
        ws['E20'] = altura
        ws['I20'] = espessura

    # Checkboxes fixos
    ws['J20'] = "X"  # FUNDO
    ws['B38'] = "X"  # NÃO SANFONA
    ws['G40'] = "X"  # QUADRADO

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# Interface Streamlit
st.title("✅ Sistema Simplificado de Preenchimento de Fichas - SACOS")
pdf = st.file_uploader("Envie o PDF do modelo SACOS", type=["pdf"])

if pdf:
    try:
        dados = extrair_dados(pdf)
        planilha = preencher_saco(dados)
        st.download_button("⬇️ Baixar Ficha Preenchida", planilha, "FICHA_PREENCHIDA.xlsx")
        st.success("Pronto! Planilha gerada com sucesso.")
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")

    if not re.search(r"(\d+)X(\d+)X([\d,]+)", dados):
        st.warning("⚠️ Atenção: Medidas não foram encontradas no PDF enviado.")
