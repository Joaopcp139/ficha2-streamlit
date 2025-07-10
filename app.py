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
    """Preenche células mescladas de forma segura"""
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and str(texto_busca).strip() in str(cell.value).strip():
                try:
                    # Verifica se a célula à direita está mesclada
                    for merged_range in ws.merged_cells.ranges:
                        if (cell.row, cell.column + 1) in merged_range:
                            # Preenche a primeira célula do range mesclado
                            ws.cell(row=merged_range.min_row, 
                                   column=merged_range.min_col, 
                                   value=valor)
                            return True
                    
                    # Se não está mesclada, preenche normalmente
                    ws.cell(row=cell.row, column=cell.column + 1, value=valor)
                    return True
                except Exception as e:
                    st.warning(f"Célula {cell.coordinate} não pode ser preenchida: {str(e)}")
                    continue
    return False

def preencher_planilha_saco(ws, dados):
    """Preenche a planilha SACO.xlsx com tratamento especial para células mescladas"""
    try:
        # Preenche campos básicos
        preencher_celula_mesclada(ws, "1.1 CLIENTE:", dados.get("cliente", ""))
        preencher_celula_mesclada(ws, "1.3 PRODUTO:", dados.get("produto", ""))
        preencher_celula_mesclada(ws, "1.2 CÓD. PRODUTO:", dados.get("codigo", ""))
        
        # Preenche medidas
        preencher_celula_mesclada(ws, "2.3 LARGURA", dados.get("largura", ""))
        preencher_celula_mesclada(ws, "2.4 COMPRIMENTO", dados.get("comprimento", ""))
        preencher_celula_mesclada(ws, "2.5 ESPESSURA", dados.get("espessura", ""))
        preencher_celula_mesclada(ws, "2.6 LARGURA", dados.get("largura", ""))
        preencher_celula_mesclada(ws, "2.7 COMPRIMENTO", dados.get("comprimento", ""))
        preencher_celula_mesclada(ws, "2.8 ESPESSURA", dados.get("espessura", ""))
        
        # Preenche quantidade e observações
        preencher_celula_mesclada(ws, "QTDE DE SACOS POR AMARRAÇÃO", dados.get("qtd_sacos", ""))
        preencher_celula_mesclada(ws, "OBSERVAÇÕES", dados.get("observacoes", ""))
        
        # Marca checkboxes conforme exemplo
        ws['J20'].value = "X"  # FUNDO (SOLDA)
        ws['B38'].value = "X"  # NÃO (SANFONA)
        ws['G40'].value = "X"  # QUADRADO (FUNDO)
        
        return True
    except Exception as e:
        st.error(f"Erro ao preencher planilha: {str(e)}")
        return False

def processar_pdf(texto):
    """Extrai dados do PDF e retorna um dicionário"""
    dados = {}
    
    def extrair(padrao):
        match = re.search(padrao, texto, re.IGNORECASE)
        return match.group(1).strip() if match else ""
    
    dados["cliente"] = extrair(r"NOME DO CLIENTE[:\s]*(.*)")
    dados["produto"] = extrair(r"PRODUTO[:\s]*(.*)")
    dados["codigo"] = extrair(r"PEDIDO N[º°:\s]*(.*)")
    dados["largura"] = extrair(r"LARGURA\s*\(mm\)[:\s]*(\d+)")
    dados["comprimento"] = extrair(r"COMPRIMENTO[:\s]*(\d+)")
    espessura = extrair(r"ESPESSURA\s*\(p/ parede\)[:\s]*([\d,]+)")
    dados["espessura"] = espessura.replace(",", ".") if espessura else ""
    dados["qtd_sacos"] = extrair(r"OTDE DE SACOS P/ PACOTE[:\s]*(\d+)")
    dados["observacoes"] = extrair(r"OBSERVAÇÕES[\s\n]*(.*?)(?=\n\s*\n|$)")
    
    return dados

# Interface principal
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
