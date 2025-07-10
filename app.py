import streamlit as st
import PyPDF2
import re
import pandas as pd
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

def preencher_planilha(modelo, dados_extraidos):
    try:
        if modelo == "saco":
            planilha_path = "SACO.xlsx"
        else:
            planilha_path = "FILME.xlsx"

        wb = load_workbook(planilha_path)
        ws = wb.active

        # Padrões de busca para os dados
        padroes = {
            "cliente": r"(?:CLIENTE|CLIENTES)[:\s]*(.*)",
            "produto": r"PRODUTO[:\s]*(.*)",
            "cod_produto": r"(?:CÓD\. PRODUTO|PEDIDO N°?|PEDIDO)[:\s]*(.*)",
            "largura": r"LARGURA[ (A-Z)]*[:\s]*(\d+[,.]?\d*)",
            "comprimento": r"COMPRIMENTO[ (A-Z)]*[:\s]*(\d+[,.]?\d*)",
            "espessura": r"ESPESSURA[ (A-Z)]*[:\s]*(\d+[,.]?\d*)",
            "passo": r"PASSO[ (A-Z)]*[:\s]*(\d+[,.]?\d*)",
            "observacoes": r"OBSERVAÇÕES[\s\n]*(.*?)(?=\n\s*\n|$)"
        }

        # Extrair dados usando os padrões
        dados = {}
        for chave, padrao in padroes.items():
            match = re.search(padrao, dados_extraidos, re.IGNORECASE)
            if match:
                dados[chave] = match.group(1).strip()

        # Preencher a planilha
        for row in ws.iter_rows():
            for cell in row:
                if cell.value:
                    valor_celula = str(cell.value).strip()
                    
                    # Preencher cliente
                    if "1.1 CLIENTE:" in valor_celula and "cliente" in dados:
                        ws.cell(row=cell.row, column=cell.column+1, value=dados["cliente"])
                    
                    # Preencher produto
                    elif "1.3 PRODUTO:" in valor_celula and "produto" in dados:
                        ws.cell(row=cell.row, column=cell.column+1, value=dados["produto"])
                    
                    # Preencher código do produto
                    elif "1.2 CÓD. PRODUTO:" in valor_celula and "cod_produto" in dados:
                        ws.cell(row=cell.row, column=cell.column+1, value=dados["cod_produto"])
                    
                    # Preencher largura (SACO e FILME)
                    elif "2.3 LARGURA" in valor_celula and "largura" in dados:
                        ws.cell(row=cell.row, column=cell.column+1, value=float(dados["largura"].replace(",", ".")))
                    
                    # Preencher comprimento (SACO)
                    elif "2.4 COMPRIMENTO" in valor_celula and "comprimento" in dados:
                        ws.cell(row=cell.row, column=cell.column+1, value=float(dados["comprimento"].replace(",", ".")))
                    
                    # Preencher espessura
                    elif "2.5 ESPESSURA" in valor_celula and "espessura" in dados:
                        ws.cell(row=cell.row, column=cell.column+1, value=float(dados["espessura"].replace(",", ".")))
                    
                    # Preencher passo (FILME)
                    elif "2.4 PASSO DA FOTOCÉLULA" in valor_celula and "passo" in dados:
                        ws.cell(row=cell.row, column=cell.column+1, value=float(dados["passo"].replace(",", ".")))
                    
                    # Preencher observações
                    elif "OBSERVAÇÕES" in valor_celula and "observacoes" in dados:
                        ws.cell(row=cell.row, column=cell.column+1, value=dados["observacoes"])

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    except Exception as e:
        st.error(f"Erro ao preencher planilha: {str(e)}")
        return None

# Interface do Streamlit
st.title("📋 Gerador de Fichas Técnicas")
st.markdown("**SACOS e FILMES**")

uploaded_pdf = st.file_uploader("Envie o PDF da ficha", type=["pdf"])

if uploaded_pdf:
    with st.spinner("Processando PDF..."):
        texto = extrair_dados_pdf(uploaded_pdf)
        
        if texto:
            modelo = identificar_modelo(texto)
            st.success(f"✅ Modelo detectado: {modelo.upper()}")
            
            with st.expander("Visualizar texto extraído"):
                st.text(texto)
            
            planilha_preenchida = preencher_planilha(modelo, texto)
            
            if planilha_preenchida:
                st.download_button(
                    label="📥 Baixar planilha preenchida",
                    data=planilha_preenchida,
                    file_name=f"FICHA_{modelo.upper()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )