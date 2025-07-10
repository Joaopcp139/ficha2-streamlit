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

def preencher_celula_segura(ws, texto_busca, valor, col_offset=1):
    """Versão aprimorada que lida com todos os tipos de células"""
    for row in ws.iter_rows():
        for cell in row:
            try:
                # Extrai o valor da célula de forma segura
                cell_value = str(cell.value) if cell.value is not None else ""
                
                # Verifica se o texto_busca está no valor da célula
                if str(texto_busca).strip() in cell_value.strip():
                    try:
                        # Verifica se a célula de destino está mesclada
                        target_row = cell.row
                        target_col = cell.column + col_offset
                        
                        # Encontra o intervalo mesclado que contém a célula alvo
                        for merged_range in ws.merged_cells.ranges:
                            if (target_row, target_col) in merged_range:
                                # Preenche a célula principal do intervalo mesclado
                                ws.cell(
                                    row=merged_range.min_row,
                                    column=merged_range.min_col,
                                    value=str(valor) if valor is not None else ""
                                )
                                return True
                        
                        # Se não está mesclada, preenche normalmente
                        ws.cell(
                            row=target_row,
                            column=target_col,
                            value=str(valor) if valor is not None else ""
                        )
                        return True
                        
                    except Exception as e:
                        st.warning(f"Célula {get_column_letter(cell.column)}{cell.row} não pôde ser preenchida. Erro: {str(e)}")
                        continue
                        
            except Exception as e:
                # Se falhar ao ler a célula, apenas continua
                continue
                
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
        dados = {}
        for chave, padroes in campos.items():
            for padrao in padroes:
                match = re.search(padrao, texto_extraido, re.IGNORECASE)
                if match:
                    dados[chave] = match.group(1).strip()
                    break

        # Mostra os dados encontrados para debug
        with st.expander("Dados extraídos do PDF"):
            st.json(dados)

        # Preenche a planilha de acordo com o modelo
        if modelo == "saco":
            mapeamento = [
                ("1.1 CLIENTE:", "cliente"),
                ("1.3 PRODUTO:", "produto"),
                ("1.2 CÓD. PRODUTO:", "cod_produto"),
                ("2.3 LARGURA", "largura"),
                ("2.4 COMPRIMENTO", "comprimento"),
                ("2.5 ESPESSURA", "espessura"),
                ("2.6 LARGURA", "largura_final"),
                ("2.7 COMPRIMENTO", "comprimento"),
                ("2.8 ESPESSURA", "espessura_final"),
                ("OBSERVAÇÕES", "observacoes"),
                ("QTDE DE SACOS POR AMARRAÇÃO", "qtd_sacos")
            ]
        else:  # filme
            mapeamento = [
                ("1.1 CLIENTE:", "cliente"),
                ("1.3 PRODUTO:", "produto"),
                ("1.2 CÓD. PRODUTO:", "cod_produto"),
                ("2.3 LARGURA", "largura"),
                ("2.4 PASSO DA FOTOCÉLULA", "passo"),
                ("2.5 ESPESSURA", "espessura"),
                ("OBSERVAÇÕES", "observacoes"),
                ("2.9 PESO POR BOBINA:", "peso_bobina")
            ]
        
        for texto_celula, chave_dado in mapeamento:
            valor = dados.get(chave_dado, "")
            if not preencher_celula_segura(ws, texto_celula, valor):
                st.warning(f"Campo '{texto_celula}' não pode ser preenchido")

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
