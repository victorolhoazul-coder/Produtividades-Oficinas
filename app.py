elif modulo_selecionado == "💶 Faturação e Orçamento Diário":
        st.subheader(f"Faturação e Orçamento Diário - {meses_pt[mes_selecionado].capitalize()} {ano_corrente}")
        
        # 1. Carregamento de Orçamentos
        excel_orc_path = "Orçamento-agrupamento oficinas.xlsx"
        df_orc = pd.DataFrame()
        df_agrup = pd.DataFrame()
        
        target_orc_file = file_orc_up if file_orc_up is not None else (excel_orc_path if os.path.exists(excel_orc_path) else None)
        
        if target_orc_file is not None:
            try:
                df_orc = pd.read_excel(target_orc_file, sheet_name="Oficina_ORÇAMENTO")
                df_agrup = pd.read_excel(target_orc_file, sheet_name="Agrupamento_oficinas")
            except Exception as e:
                st.error(f"Erro ao ler o ficheiro de orçamentos: {e}")
        
        # 2. Leitura com header=5 para apanhar exatamente a linha dos cabeçalhos do CSV
        df_fact_total = pd.DataFrame()
        csv_fontes = [file_fact1_up, file_fact2_up]
        
        if all(f is None for f in csv_fontes):
            csv_fontes = glob.glob(f"*FACTURAÇÃO*{ano_corrente}_{mes_selecionado:02d}.csv")
            if not csv_fontes:
                csv_fontes = glob.glob("*FACTURAÇÃO*.csv")
            if not csv_fontes:
                csv_fontes = glob.glob("*.csv")
                
        for file_f in csv_fontes:
            if file_f is not None:
                try:
                    if hasattr(file_f, "getvalue"):
                        file_bytes = io.BytesIO(file_f.getvalue())
                    elif isinstance(file_f, str):
                        with open(file_f, "rb") as f_in:
                            file_bytes = io.BytesIO(f_in.read())
                    else:
                        file_bytes = file_f

                    t_temp = pd.DataFrame()
                    file_bytes.seek(0)
                    try:
                        temp_df = pd.read_csv(file_bytes, sep=None, engine="python", encoding="iso-8859-1", header=5, on_bad_lines='skip')
                        if len(temp_df.columns) > 1:
                            t_temp = temp_df
                    except Exception as ex_read:
                        pass
                    
                    if not t_temp.empty:
                        df_fact_total = pd.concat([df_fact_total, t_temp], ignore_index=True)
                except Exception:
                    pass
                
        if df_fact_total.empty:
            st.warning("⚠️ Não foram encontrados mapas de faturação válidos ou os ficheiros estão vazios.")
        else:
            df_fact_total.columns = [str(c).strip() for c in df_fact_total.columns]
            
            col_ofic = next((c for c in df_fact_total.columns if "OFICINA" in c.upper()), df_fact_total.columns[0])
            col_val = next((c for c in df_fact_total.columns if "TOTAL" in c.upper() or "LIQ" in c.upper()), df_fact_total.columns[24] if len(df_fact_total.columns) > 24 else df_fact_total.columns[-1])
            col_data = next((c for c in df_fact_total.columns if "DATA" in c.upper()), df_fact_total.columns[28] if len(df_fact_total.columns) > 28 else df_fact_total.columns[-2])

            df_fact_total["Oficina_CSV"] = df_fact_total[col_ofic].astype(str).str.strip()
            
            # Limpeza de linhas inválidas
            invalidos = ["NAN", "", "TOTAL", "MAPA DE FACTURAÇÃO", "None"]
            df_fact_total = df_fact_total[~df_fact_total["Oficina_CSV"].str.upper().isin(invalidos)]
            df_fact_total = df_fact_total[~df_fact_total["Oficina_CSV"].str.startswith("*")]

            # Conversão da Data
            col_data_raw = df_fact_total[col_data]
            df_fact_total["Data_Parsed"] = pd.to_datetime(col_data_raw, dayfirst=True, errors="coerce")

            # Tratamento dos valores numéricos
            col_valor_raw = df_fact_total[col_val]
            if col_valor_raw.dtype == object:
                s_clean = col_valor_raw.astype(str).str.replace("€", "", regex=False).str.strip()
                s_clean = s_clean.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
                df_fact_total["Valor_Num"] = pd.to_numeric(s_clean, errors="coerce").fillna(0.0)
            else:
                df_fact_total["Valor_Num"] = pd.to_numeric(col_valor_raw, errors="coerce").fillna(0.0)

            # 3. Mapeamento e Agrupamento
            if not df_agrup.empty:
                df_agrup.columns = [str(c).strip() for c in df_agrup.columns]
                col_agrup_oficina = df_agrup.columns[0]
                col_agrup_destino = df_agrup.columns[3] if len(df_agrup.columns) > 3 else df_agrup.columns[1]
                
                dict_agrupamento = dict(zip(df_agrup[col_agrup_oficina].astype(str).str.strip().str.upper(), df_agrup[col_agrup_destino].astype(str).str.strip()))
                df_fact_total["Oficina_Agrupada"] = df_fact_total["Oficina_CSV"].str.upper().map(dict_agrupamento).fillna(df_fact_total["Oficina_CSV"])
            else:
                df_fact_total["Oficina_Agrupada"] = df_fact_total["Oficina_CSV"]

            oficinas_disponiveis_fat = sorted([o for o in df_fact_total["Oficina_Agrupada"].dropna().astype(str).unique() if o.strip() and "MAPA" not in o.upper() and not set(o.strip()).issubset({"*"})])
            if not df_orc.empty:
                df_orc.columns = [str(c).strip() for c in df_orc.columns]
                oficinas_orc_unicas = sorted([o for o in df_orc[df_orc.columns[0]].dropna().astype(str).unique() if o.strip() and "MAPA" not in o.upper() and not set(o.strip()).issubset({"*"})])
                oficinas_disponiveis_fat = sorted(list(set(oficinas_disponiveis_fat + oficinas_orc_unicas)))
            
            if not oficinas_disponiveis_fat:
                oficinas_disponiveis_fat = ["Geral"]

            col_sel_ofic, col_empty = st.columns([2, 2])
            with col_sel_ofic:
                oficina_selecionada_fat = st.selectbox("Selecionar Oficina (Faturação):", options=oficinas_disponiveis_fat)
                
            def limpar_str(s):
                if pd.isna(s): return ""
                s_nfkd = unicodedata.normalize('NFKD', str(s))
                return "".join([c for c in s_nfkd if not unicodedata.combining(c)]).upper().strip()

            sel_clean = limpar_str(oficina_selecionada_fat)
            
            # Filtro robusto: se a oficina selecionada estiver vazia ou for "Geral", mostra tudo; caso contrário filtra por correspondência parcial limpa
            if sel_clean and sel_clean != "GERAL":
                # Dividir em palavras-chave para apanhar variações (ex: "COLISAO" e "BEJA")
                palavras_chave = [p for p in sel_clean.split() if len(p) > 2]
                if palavras_chave:
                    mask = df_fact_total["Oficina_Agrupada"].apply(limpar_str).apply(lambda x: all(k in x for k in palavras_chave))
                    if not mask.any():
                        mask = df_fact_total["Oficina_CSV"].apply(limpar_str).apply(lambda x: all(k in x for k in palavras_chave))
                    df_oficina_fat = df_fact_total[mask]
                else:
                    df_oficina_fat = df_fact_total[df_fact_total["Oficina_Agrupada"].apply(limpar_str) == sel_clean]
            else:
                df_oficina_fat = df_fact_total

            if df_oficina_fat.empty:
                df_oficina_fat = df_fact_total
            
            # Filtro estritamente pelo mês e ano selecionados
            df_mes_ano = df_oficina_fat[(df_oficina_fat["Data_Parsed"].dt.month == mes_selecionado) & (df_oficina_fat["Data_Parsed"].dt.year == ano_corrente)]
            if not df_mes_ano.empty:
                df_oficina_fat = df_mes_ano
            
            df_diario = df_oficina_fat.groupby(df_oficina_fat["Data_Parsed"].dt.date, as_index=False)["Valor_Num"].sum()
            df_diario.rename(columns={"Valor_Num": "Facturacao_Diaria", "Data_Parsed": "Data"}, inplace=True)
            df_diario = df_diario.dropna(subset=["Data"]).sort_values("Data")
            
            inicio_mes = date(ano_corrente, mes_selecionado, 1)
            if mes_selecionado == 12:
                fim_mes = date(ano_corrente + 1, 1, 1) - timedelta(days=1)
            else:
                fim_mes = date(ano_corrente, mes_selecionado + 1, 1) - timedelta(days=1)
                
            range_dias = pd.date_range(start=inicio_mes, end=fim_mes)
            df_completo_dias = pd.DataFrame({"Data": range_dias.date})
            df_completo_dias = pd.merge(df_completo_dias, df_diario, on="Data", how="left").fillna(0.0)
            df_completo_dias["Facturacao_Acumulada"] = df_completo_dias["Facturacao_Diaria"].cumsum()
            
            # 4. ORÇAMENTO MENSAL
            orcamento_mensal = 0.0
            if not df_orc.empty and len(df_orc.columns) >= 5:
                col_ofic_orc = df_orc.columns[0]
                col_mes_orc = df_orc.columns[3]
                col_val_orc = df_orc.columns[4]
                
                nome_mes_alvo = limpar_str(meses_pt[mes_selecionado])
                
                for _, row in df_orc.iterrows():
                    ofic_row = limpar_str(row[col_ofic_orc])
                    
                    if sel_clean in ofic_row or ofic_row in sel_clean or all(k in ofic_row for k in sel_clean.split() if len(k) > 2):
                        val_mes_cell = row[col_mes_orc]
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
                            val_raw = row[col_val_orc]
                            try:
                                if isinstance(val_raw, (int, float)):
                                    v_num = float(val_raw)
                                else:
                                    v_str = str(val_raw).replace("€", "").replace(" ", "").replace(".", "").replace(",", ".")
                                    v_num = float(v_str)
                                if v_num > 0:
                                    orcamento_mensal = v_num
                                    break
                            except:
                                pass
                                
            orcamento_diario = orcamento_mensal / dias_uteis_padrao_mes if dias_uteis_padrao_mes > 0 else 0.0
            
            orcamentos_ac = []
            acum = 0.0
            for d in df_completo_dias["Data"]:
                if d.weekday() < 5:
                    acum += orcamento_diario
                orcamentos_ac.append(acum)
            df_completo_dias["Orcamento_Acumulado"] = orcamentos_ac
            
            facturacao_a_data = df_completo_dias["Facturacao_Acumulada"].iloc[-1]
            desvio_orcamento = facturacao_a_data - orcamento_mensal
            
            col_graf, col_kpi = st.columns([3, 1])
            
            with col_graf:
                base = alt.Chart(df_completo_dias).encode(
                    x=alt.X('Data:T', title='Data', axis=alt.Axis(format='%d/%m/%Y', labelAngle=-45))
                )
                
                barras = base.mark_bar(color='#115e59').encode(
                    y=alt.Y('Facturacao_Acumulada:Q', title='Valor (€)'),
                    tooltip=['Data', 'Facturacao_Diaria', 'Facturacao_Acumulada']
                )
                
                linha_orc = base.mark_line(color='#f97316', strokeWidth=3).encode(
                    y='Orcamento_Acumulado:Q',
                    tooltip=['Data', 'Orcamento_Acumulado']
                )
                
                chart = (barras + linha_orc).properties(height=400).interactive()
                st.altair_chart(chart, use_container_width=True)
                
            with col_kpi:
                st.metric(label="ORÇAMENTO MENSAL", value=f"{orcamento_mensal:,.0f} €".replace(",", " "))
                st.metric(label="ORÇAMENTO DIÁRIO", value=f"{orcamento_diario:,.0f} €".replace(",", " "))
                st.metric(label="FATURAÇÃO À DATA", value=f"{facturacao_a_data:,.0f} €".replace(",", " "))
                st.metric(label="DESVIO ORÇAMENTO", value=f"{desvio_orcamento:,.0f} €".replace(",", " "), delta=f"{desvio_orcamento:,.0f} €".replace(",", " "))
