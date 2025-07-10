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

        # Antes de preencher, limpamos células problemáticas conhecidas
        celulas_problematicas = ['A6', 'A7', 'I6', 'A15', 'E15', 'I15', 'A19', 'E19', 'I19', 'A39']
        for coord in celulas_problematicas:
            try:
                ws[coord].value = ""  # Limpa o valor da célula
            except:
                pass

        # [Restante do código de extração e mapeamento permanece igual...]

        # Preenche a planilha com tratamento reforçado
        mapeamento_campos = {
            "saco": [
                ("1.1 CLIENTE:", "cliente"),
                ("1.3 PRODUTO:", "produto"),
                ("1.2 CÓD. PRODUTO:", "cod_produto"),
                # ... outros campos
            ],
            "filme": [
                # ... mapeamento para filmes
            ]
        }

        for texto_celula, chave_dado in mapeamento_campos[modelo]:
            valor = dados.get(chave_dado, "")
            if not preencher_celula_segura(ws, texto_celula, valor):
                st.warning(f"Campo '{texto_celula}' não pode ser preenchido")

        # [Restante do código...]
