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

def encontrar_valor(texto, padroes):
    for padrao in padroes:
        match = re.search(padrao, texto, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return ""

def preencher_celula_segura(ws, texto_busca, valor, col_offset=1):
    """Preenche células evitando células mescladas"""
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and str(texto_busca).strip() in str(cell.value).strip():
                try:
                    # Verifica se a célula alvo está mesclada
                    for merged_range in ws.merged_cells.ranges:
                        if (cell.row, cell.column + col_offset) in merged_range:
                            # Preenche a primeira célula do merge
                            ws.cell(row=merged_range.min_row, 
                                   column=merged_range.min_col, 
                                   value=valor)
                            return True
                    
                    # Se não está mesclada, preenche normalmente
                    ws.cell(row=cell.row, column=cell.column+col_offset, value=valor)
                    return True
                except Exception as e:
                    st.warning(f"Célula {get_column_letter(cell.column)}{cell.row} não pôde ser preenchida: {str(e)}")
                    return False
    return False

def preencher_planilha_avancado(modelo, texto_extraido):
    try:
        # Carrega a planilha correta
        planilha_path = "SACO.xlsx" if modelo == "saco" else "FILME.xlsx"
        wb = load_workbook(planilha_path)
        ws = wb.active

        # Dicionário de mapeamento de campos
        campos = {
            "cliente": [r"NOME DO CLIENTE[:\s]*(.*)", r"CLIENTE[:\s]*(.*)"],
            "produto": [r"PRODUTO[:\s]*(.*)"],
            "cod_produto": [r"PEDIDO N[º°:\s]*(.*)", r"O\.C\.\s*[:\s]*(.*)"],
            "largura": [r"LARGURA\s*\(mm\)[:\s]*(\d+[,.]?\d*)"],
            "largura_final": [r"LARGURA FINAL\s*\(mm\)[:\s]*(\d+[,.]?\d*)"],
            "comprimento": [r"COMPRIMENTO[:\s]*(\d+[,.]?\d*)"],
            "passo": [r"PASSO\s*\(mm\)[:\s]*(\d+[,.]?\d*)"],
            "espessura": [r"ESPESSURA\s*\(p/ parede\)[:\s]*(\d+[,.]?\d*)"],
            "espessura_final": [r"ESPESSURA FINAL[:\s]*(\d+[,.]?\d*)"],
            "peso_bobina": [r"OTDE EM KG P/ BOBINA[:\s]*(\d+)"],
            "qtd_sacos": [r"OTDE DE SACOS P/ PACOTE[:\s]*(\d+)"],
            "observacoes": [r"OBSERVAÇÕES[\s\n]*(.*?)(?=\n\s*\n|$)"]
        }

        # Extrai os dados
        dados = {chave: encontrar_valor(texto_extraido, padroes) 
                for chave, padroes in campos.items()}

        # Mostra os dados encontrados para debug
        with st.expander("Dados extraídos do PDF"):
            st.json(dados)

        # Preenche a planilha de acordo com o modelo
        if modelo == "saco":
            preencher_celula_segura(ws, "1.1 CLIENTE:", dados.get("cliente", ""))
            preencher_celula_segura(ws, "1.3 PRODUTO:", dados.get("produto", ""))
            preencher_celula_segura(ws, "1.2 CÓD. PRODUTO:", dados.get("cod_produto", ""))
            preencher_celula_segura(ws, "2.3 LARGURA", dados.get("largura", ""))
            preencher_celula_segura(ws, "2.4 COMPRIMENTO", dados.get("comprimento", ""))
            preencher_celula_segura(ws, "2.5 ESPESSURA", dados.get("espessura", ""))
            preencher_celula_segura(ws, "2.6 LARGURA", dados.get("largura_final", ""))
            preencher_celula_segura(ws, "2.7 COMPRIMENTO", dados.get("comprimento", ""))
            preencher_celula_segura(ws, "2.8 ESPESSURA", dados.get("espessura_final", ""))
        else:  # filme
            preencher_celula_segura(ws, "1.1 CLIENTE:", dados.get("cliente", ""))
            preencher_celula_segura(ws, "1.3 PRODUTO:", dados.get("produto", ""))
            preencher_celula_segura(ws, "1.2 CÓD. PRODUTO:", dados.get("cod_produto", ""))
            preencher_celula_segura(ws, "2.3 LARGURA", dados.get("largura", ""))
            preencher_celula_segura(ws, "2.4 PASSO DA FOTOCÉLULA", dados.get("passo", ""))
            preencher_celula_segura(ws, "2.5 ESPESSURA", dados.get("espessura", ""))
        
        # Campos comuns
        preencher_celula_segura(ws, "OBSERVAÇÕES", dados.get("observacoes", ""))
        preencher_celula_segura(ws, "2.9 PESO POR BOBINA:", dados.get("peso_bobina", ""))
        preencher_celula_segura(ws, "QTDE DE SACOS POR AMARRAÇÃO", dados.get("qtd_sacos", ""))

        # Salva a planilha
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    except Exception as e:
        st.error(f"Erro ao processar a planilha: {str(e)}")
        return None

# Interface do Streamlit
st.title("📋 Sistema de Fichas Técnicas Automatizado")
st.markdown("**SACOS e FILMES**")

uploaded_file = st.file_uploader("Envie o PDF da ficha técnica", type=["pdf"])

if uploaded_file:
    with st.spinner("Processando arquivo..."):
        texto = extrair_dados_pdf(uploaded_file)
        
        if texto:
            modelo = identificar_modelo(texto)
            st.success(f"✅ Tipo identificado: {modelo.upper()}")
            
            planilha = preencher_planilha_avancado(modelo, texto)
            
            if planilha:
                st.download_button(
                    label="⬇️ Baixar Ficha Técnica Preenchida",
                    data=planilha,
                    file_name=f"FICHA_{modelo.upper()}_PREENCHIDA.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success("Planilha gerada com sucesso!")
            else:
                st.error("Ocorreu um erro ao preencher a planilha.")
