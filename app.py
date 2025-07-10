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

def preencher_celula_mesclada(ws, texto_busca, valor):
    for row in ws.iter_rows():
        for cell in row:
            try:
                cell_val = str(cell.value).strip() if isinstance(cell.value, str) else ""
                if texto_busca.strip() in cell_val:
                    destino_col = cell.column + 1
                    destino_row = cell.row

                    for merged_range in ws.merged_cells.ranges:
                        if (destino_row, destino_col) in merged_range:
                            ws.cell(row=merged_range.min_row, column=merged_range.min_col, value=valor)
                            return True

                    ws.cell(row=destino_row, column=destino_col, value=valor)
                    return True
            except Exception as e:
                st.warning(f"Célula {cell.coordinate} não pode ser preenchida: {str(e)}")
    return False

def preencher_planilha_saco(ws, dados):
    try:
        preencher_celula_mesclada(ws, "1.1 CLIENTE:", dados.get("cliente", ""))
        preencher_celula_mesclada(ws, "1.3 PRODUTO:", dados.get("produto", ""))
        preencher_celula_mesclada(ws, "1.2 CÓD. PRODUTO:", dados.get("codigo", ""))
        preencher_celula_mesclada(ws, "2.3 LARGURA", dados.get("largura", ""))
        preencher_celula_mesclada(ws, "2.4 COMPRIMENTO", dados.get("comprimento", ""))
        preencher_celula_mesclada(ws, "2.5 ESPESSURA", dados.get("espessura", ""))
        preencher_celula_mesclada(ws, "2.6 LARGURA", dados.get("largura", ""))
        preencher_celula_mesclada(ws, "2.7 COMPRIMENTO", dados.get("comprimento", ""))
        preencher_celula_mesclada(ws, "2.8 ESPESSURA", dados.get("espessura", ""))
        preencher_celula_mesclada(ws, "QTDE DE SACOS POR AMARRAÇÃO", dados.get("qtd_sacos", ""))
        preencher_celula_mesclada(ws, "OBSERVAÇÕES", dados.get("observacoes", ""))
        if dados.get("fundo") == "SIM":
            ws['J20'].value = "X"
        if dados.get("sanfona") == "NÃO":
            ws['B38'].value = "X"
        ws['G40'].value = "X"  # QUADRADO (sempre marcado como exemplo)
        return True
    except Exception as e:
        st.error(f"Erro ao preencher planilha: {str(e)}")
        return False

def processar_pdf(texto):
    dados = {}
    def extrair(padrao):
        match = re.search(padrao, texto, re.IGNORECASE)
        return match.group(1).strip() if match else ""

    dados["cliente"] = extrair(r"CLIENTES:\s*\d+\s*-\s*(.*)")
    dados["produto"] = extrair(r"PRODUTO:\s*(.*?)(?=QTDE|LARGURA|$)").split("-")[-1].strip()
    dados["codigo"] = extrair(r"PEDIDO N[:º\s]*(\d+)")
    dados["largura"] = extrair(r"LARGURA:\s*(\d+)")
    dados["comprimento"] = extrair(r"PASSO:\s*(\d+)")
    espessura_final = extrair(r"ESPESSURA FINAL[:\s]*(0[,\.]\d+)")
    dados["espessura"] = espessura_final.replace(",", ".") if espessura_final else ""
    dados["qtd_sacos"] = extrair(r"quant de pacotes[:\s]*(\d+)")
    dados["observacoes"] = extrair(r"OBSERVAÇÕES[:\s\n]*(.*?)\n") or extrair(r"OUTROS[:\s]*(\w+)")
    dados["sanfona"] = "NÃO" if re.search(r"SANFONA SIM:\s*Off.*?SANFONA NAO:\s*Yes", texto, re.IGNORECASE | re.DOTALL) else "SIM"
    dados["fundo"] = "SIM" if re.search(r"FUNDO:\s*Yes", texto) else "NÃO"

    return dados

st.title("📋 Sistema Automático de Fichas Técnicas")
uploaded_file = st.file_uploader("Envie o PDF da ficha", type=["pdf"])

if uploaded_file:
    texto = extrair_dados_pdf(uploaded_file)
    if texto:
        modelo = identificar_modelo(texto)
        st.success(f"Modelo detectado: {modelo.upper()}")
        dados = processar_pdf(texto)
        with st.expander("Ver dados extraídos"):
            st.json(dados)
        try:
            wb = load_workbook("SACO.xlsx")
            ws = wb.active
            if preencher_planilha_saco(ws, dados):
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                st.download_button(
                    label="⬇️ Baixar Ficha Técnica Preenchida",
                    data=output,
                    file_name=f"FICHA_{modelo.upper()}_PREENCHIDA.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success("Planilha gerada com sucesso!")
            else:
                st.error("Ocorreu um erro ao preencher a planilha")
        except Exception as e:
            st.error(f"Erro ao processar a planilha: {str(e)}")
