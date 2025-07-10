import streamlit as st
import PyPDF2
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

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

def extrair_valor_celula(cell):
    """Extrai valor de célula mesmo se for tupla ou formato complexo"""
    try:
        if cell.value is None:
            return ""
        if isinstance(cell.value, (str, int, float)):
            return str(cell.value)
        if isinstance(cell.value, tuple):
            return " ".join(str(x) for x in cell.value if x)
        return str(cell.value)
    except:
        return ""

def preencher_celula_segura(ws, texto_busca, valor, col_offset=1):
    """Versão robusta que ignora células problemáticas"""
    texto_busca = str(texto_busca).strip()
    valor = str(valor) if valor is not None else ""
    
    for row in ws.iter_rows():
        for cell in row:
            try:
                cell_value = extrair_valor_celula(cell)
                if texto_busca in cell_value:
                    target_col = cell.column + col_offset
                    target_cell = ws.cell(row=cell.row, column=target_col)
                    
                    # Verifica se a célula alvo está mesclada
                    for merged_range in ws.merged_cells.ranges:
                        if (cell.row, target_col) in merged_range:
                            target_cell = ws.cell(row=merged_range.min_row, 
                                                column=merged_range.min_col)
                            break
                    
                    # Tenta preencher a célula
                    try:
                        target_cell.value = valor
                        return True
                    except:
                        continue
            except:
                continue
    return False

def preencher_planilha(modelo, texto_extraido):
    try:
        planilha_path = "SACO.xlsx" if modelo == "saco" else "FILME.xlsx"
        wb = load_workbook(planilha_path)
        ws = wb.active

        # Padrões de extração
        padroes = {
            "cliente": r"(?:CLIENTE|NOME DO CLIENTE)[:\s]*(.*?)(?:\n|$)",
            "produto": r"PRODUTO[:\s]*(.*?)(?:\n|$)", 
            "codigo": r"(?:CÓD\. PRODUTO|PEDIDO N°?|O\.C\.)[:\s]*(.*?)(?:\n|$)",
            "largura": r"LARGURA\s*\(mm\)[:\s]*(\d+[,.]?\d*)",
            "comprimento": r"COMPRIMENTO[:\s]*(\d+[,.]?\d*)",
            "espessura": r"ESPESSURA\s*\(.*\)[:\s]*(\d+[,.]?\d*)",
            "observacoes": r"OBSERVAÇÕES[\s\n]*(.*?)(?=\n\s*\n|$)"
        }

        # Extrai dados
        dados = {}
        for campo, regex in padroes.items():
            match = re.search(regex, texto_extraido, re.IGNORECASE)
            if match:
                dados[campo] = match.group(1).strip()

        # Mapeamento de campos
        campos = {
            "saco": [
                ("1.1 CLIENTE:", "cliente"),
                ("1.3 PRODUTO:", "produto"),
                ("1.2 CÓD. PRODUTO:", "codigo"),
                ("2.3 LARGURA", "largura"),
                ("2.4 COMPRIMENTO", "comprimento"),
                ("2.5 ESPESSURA", "espessura"),
                ("OBSERVAÇÕES", "observacoes")
            ],
            "filme": [
                ("1.1 CLIENTE:", "cliente"),
                ("1.3 PRODUTO:", "produto"), 
                ("1.2 CÓD. PRODUTO:", "codigo"),
                ("2.3 LARGURA", "largura"),
                ("2.4 PASSO DA FOTOCÉLULA", "comprimento"),
                ("2.5 ESPESSURA", "espessura"),
                ("OBSERVAÇÕES", "observacoes")
            ]
        }

        # Preenche planilha
        for texto_celula, chave in campos[modelo]:
            valor = dados.get(chave, "")
            preencher_celula_segura(ws, texto_celula, valor)

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    except Exception as e:
        st.error(f"Erro ao processar planilha: {str(e)}")
        return None

# Interface
st.title("📋 Gerador de Fichas Técnicas")

uploaded_file = st.file_uploader("Envie o PDF da ficha", type=["pdf"])

if uploaded_file:
    texto = extrair_dados_pdf(uploaded_file)
    if texto:
        modelo = identificar_modelo(texto)
        st.success(f"Modelo detectado: {modelo.upper()}")
        
        with st.expander("Ver texto extraído"):
            st.text(texto[:1000] + ("..." if len(texto) > 1000 else ""))
        
        planilha = preencher_planilha(modelo, texto)
        
        if planilha:
            st.download_button(
                label="⬇️ Baixar Planilha",
                data=planilha,
                file_name=f"FICHA_{modelo.upper()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Planilha gerada com sucesso!")
        else:
            st.error("Falha ao gerar planilha")
