def tratar_valor(v):
    if isinstance(v, tuple):
        return str(v[0]) if v else ""
    if v is None:
        return ""
    return str(v)

def preencher_celula_mesclada(ws, texto_busca, valor):
    valor = tratar_valor(valor)

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
        ws['G40'].value = "X"
        return True
    except Exception as e:
        st.error(f"Erro ao preencher planilha: {str(e)}")
        return False
