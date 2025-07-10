def preencher_planilha_avancado(modelo, texto_extraido):
    try:
        # ... (código anterior)
        
        # Preenche a planilha com tratamento adicional
        if modelo == "saco":
            campos_para_preencher = [
                ("1.1 CLIENTE:", dados.get("cliente", "")),
                ("1.3 PRODUTO:", dados.get("produto", "")),
                # ... (todos os outros campos)
            ]
        else:  # filme
            campos_para_preencher = [
                ("1.1 CLIENTE:", dados.get("cliente", "")),
                # ... (todos os outros campos)
            ]
            
        for texto_celula, valor in campos_para_preencher:
            if not preencher_celula_segura(ws, texto_celula, valor):
                st.warning(f"Campo '{texto_celula}' não encontrado na planilha")
        
        # ... (restante do código)
