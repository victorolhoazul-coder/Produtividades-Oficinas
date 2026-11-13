# 2. Leitura blindada e direta dos CSVs carregados via sidebar ou disco
        df_fact_total = pd.DataFrame()
        csv_fontes = [file_fact1_up, file_fact2_up]
        
        # Recolher também do disco caso estejam na pasta
        arquivos_disco = glob.glob("*MOTOREX*.csv") + glob.glob("*MULTIAUTO*.csv") + glob.glob("*MAPA_FACTURAÇÃO*.csv") + glob.glob("*.csv")
        
        lista_streams = []
        for f in csv_fontes:
            if f is not None:
                lista_streams.append(f)
        
        if not lista_streams:
            for arq in arquivos_disco:
                if os.path.exists(arq):
                    lista_streams.append(arq)

        for file_f in lista_streams:
            try:
                if hasattr(file_f, "getvalue"):
                    file_bytes = io.BytesIO(file_f.getvalue())
                elif isinstance(file_f, str):
                    with open(file_f, "rb") as f_in:
                        file_bytes = io.BytesIO(f_in.read())
                else:
                    file_bytes = file_f

                # Tenta ler ignorando cabeçalhos mal formatados
                for sep_val in [';', ',', '\t']:
                    try:
                        file_bytes.seek(0)
                        temp_df = pd.read_csv(file_bytes, sep=sep_val, engine="python", encoding="iso-8859-1", header=None, on_bad_lines='skip')
                        if not temp_df.empty and len(temp_df.columns) > 2:
                            df_fact_total = pd.concat([df_fact_total, temp_df], ignore_index=True)
                            break
                    except Exception:
                        pass
            except Exception:
                pass
                
        if df_fact_total.empty:
            st.warning("⚠️ Não foram encontrados mapas de faturação válidos ou os ficheiros estão vazios.")
        else:
            # Identifica automaticamente as colunas úteis (Oficina na col 0, Data e Valor nas restantes)
            col_ofic_idx = 0
            col_data_idx = min(28, len(df_fact_total.columns) - 1)
            col_val_idx = min(24, len(df_fact_total.columns) - 1)

            # Procura refinada por colunas que contenham datas válidas
            for col in df_fact_total.columns:
                parsed_dates = pd.to_datetime(df_fact_total[col], dayfirst=True, errors='coerce')
                if parsed_dates.notna().mean() > 0.05:
                    col_data_idx = col
                    break

            # Procura refinada por valores numéricos
            for col in reversed(df_fact_total.columns):
                if col != col_data_idx:
                    sample_vals = df_fact_total[col].astype(str).str.replace('€', '').str.strip()
                    numeric_check = pd.to_numeric(sample_vals.str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce')
                    if numeric_check.notna().mean() > 0.05:
                        col_val_idx = col
                        break

            df_fact_total["Oficina_CSV"] = df_fact_total[col_ofic_idx].astype(str).str.strip()
            
            invalidos = ["NAN", "", "TOTAL", "MAPA DE FACTURAÇÃO", "None", "nan", "SUBTOTAL", "OFICINA"]
            df_fact_total = df_fact_total[~df_fact_total["Oficina_CSV"].str.upper().isin(invalidos)]
            df_fact_total = df_fact_total[~df_fact_total["Oficina_CSV"].str.startswith("*")]

            col_data_raw = df_fact_total[col_data_idx]
            df_fact_total["Data_Parsed"] = pd.to_datetime(col_data_raw, dayfirst=True, errors="coerce")

            col_valor_raw = df_fact_total[col_val_idx]
            s_clean = col_valor_raw.astype(str).str.replace("€", "", regex=False).str.strip()
            s_clean = s_clean.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
            df_fact_total["Valor_Num"] = pd.to_numeric(s_clean, errors="coerce").fillna(0.0)
