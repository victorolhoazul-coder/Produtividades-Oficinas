orcamento_mensal = 0.0
            if not df_orc.empty and len(df_orc.columns) >= 5:
                col_ofic = df_orc.columns[0]
                col_mes = df_orc.columns[3]
                col_val = df_orc.columns[4]
                
                nome_mes_alvo = limpar_str(meses_pt[mes_selecionado])
                
                # Identificar a palavra-chave exata da cidade (ex: FARO, BEJA, EVORA, etc.)
                cidades_keywords = ["BEJA", "EVORA", "ÉVORA", "FARO", "SETUBAL", "SETÚBAL", "PORTIMAO", "PORTIMÃO", "SINES", "ALMADA"]
                cidade_alvo = next((c for c in cidades_keywords if c in sel_clean), "")
                
                for _, row in df_orc.iterrows():
                    ofic_row = limpar_str(row[col_ofic])
                    
                    if not ofic_row or "TOTAL" in ofic_row or "GRUPO" in ofic_row or "MAPA" in ofic_row:
                        continue
                    
                    # Verificação estrita: a linha do excel tem de corresponder exatamente à cidade ou à oficina selecionada
                    match_orc = False
                    if cidade_alvo and cidade_alvo in ofic_row.split():
                        match_orc = True
                    elif sel_clean == ofic_row:
                        match_orc = True
                        
                    if match_orc:
                        val_mes_cell = row[col_mes]
                        mes_encontrado = False
                        
                        if isinstance(val_mes_cell, (pd.Timestamp, datetime, date)):
                            if val_mes_cell.month == mes_selecionado:
                                mes_encontrado = True
                        else:
                            mes_row_str = limpar_str(val_mes_cell)
                            if nome_mes_alvo in mes_row_str or mes_row_str in nome_mes_alvo:
                                mes_encontrado = True
                            else:
                                try:
                                    if int(val_mes_cell) == mes_selecionado:
                                        mes_encontrado = True
                                except:
                                    pass
                        
                        if mes_encontrado:
                            val_raw = row[col_val]
                            orc_val = converter_valor_pt(val_raw)
                            if orc_val > 0:
                                orcamento_mensal = orc_val
                                break
