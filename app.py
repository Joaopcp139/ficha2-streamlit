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
    return "saco"

def preencher_planilha(modelo, texto):
    try:
        # Carrega a planilha modelo
        planilha_path = "SACO.xlsx" if modelo == "saco" else "FILME.xlsx"
        wb = load_workbook(planilha_path)
        ws = wb.active

        # Extrai dados específicos do PDF
        def extrair_valor(padrao):
            match = re.search(padrao, texto, re.IGNORECASE)
            return match.group(1).strip() if match else ""

        cliente = extrair_valor(r"NOME DO CLIENTE[:\s]*(.*)")
        produto = extrair_valor(r"PRODUTO[:\s]*(.*)")
        codigo = extrair_valor(r"PEDIDO N[º°:\s]*(.*)")
        largura = extrair_valor(r"LARGURA\s*\(mm\)[:\s]*(\d+)")
        comprimento = extrair_valor(r"COMPRIMENTO[:\s]*(\d+)")
        espessura = extrair_valor(r"ESPESSURA\s*\(p/ parede\)[:\s]*([\d,]+)")
        qtd_sacos = extrair_valor(r"OTDE DE SACOS P/ PACOTE[:\s]*(\d+)")
        observacoes = extrair_valor(r"OBSERVAÇÕES[\s\n]*(.*?)(?=\n\s*\n|$)")

        # Converte valores numéricos
        try:
            espessura = espessura.replace(",", ".")
        except:
            pass

        # Preenche os campos na planilha
        campos = {
            "1.1 CLIENTE:": cliente,
            "1.3 PRODUTO:": produto,
            "1.2 CÓD. PRODUTO:": codigo,
            "2.3 LARGURA": largura,
            "2.4 COMPRIMENTO": comprimento,
            "2.5 ESPESSURA": espessura,
            "2.6 LARGURA": largura,
            "2.7 COMPRIMENTO": comprimento, 
            "2.8 ESPESSURA": espessura,
            "QTDE DE SACOS POR AMARRAÇÃO": qtd_sacos,
            "OBSERVAÇÕES": observacoes
        }

        # Preenchimento seguro
        for row in ws.iter_rows():
            for cell in row:
                if cell.value in campos:
                    try:
                        ws.cell(row=cell.row, column=cell.column+1, value=campos[cell.value])
                    except:
                        continue

        # Configura checkboxes conforme o exemplo
        ws['J20'].value = "X"  # FUNDO (SOLDA)
        ws['B38'].value = "X"  # NÃO (SANFONA)
        ws['G40'].value = "X"  # QUADRADO (FUNDO)

        # Salva a planilha
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    except Exception as e:
        st.error(f"Erro ao preencher planilha: {str(e)}")
        return None

# Interface
st.title("📋 Preenchimento Automático de Fichas Técnicas")

uploaded_file = st.file_uploader("Envie o PDF da ficha", type=["pdf"])

if uploaded_file:
    texto = extrair_dados_pdf(uploaded_file)
    if texto:
        modelo = identificar_modelo(texto)
        st.success(f"Modelo detectado: {modelo.upper()}")
        
        planilha = preencher_planilha(modelo, texto)
        
        if planilha:
            st.download_button(
                label="⬇️ Baixar Ficha Preenchida",
                data=planilha,
                file_name=f"FICHA_{modelo.upper()}_PREENCHIDA.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Planilha gerada com sucesso!")
        else:
            st.error("Erro ao gerar planilha")
