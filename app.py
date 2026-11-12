# PROCESSAMENTO DE AUSÊNCIAS
    df_ausencias = pd.DataFrame()
    if file_ausencias_up is not None:
        df_ausencias = pd.read_excel(file_ausencias_up)
    elif path_ausencias and isinstance(path_ausencias, str) and os.path.exists(path_ausencias):
        df_ausencias = pd.read_excel(path_ausencias)

    if not df_ausencias.empty:
        col_desc_aus = df_ausencias.columns[1] if len(df_ausencias.columns) > 1 else None
        col_op_aus = next((c for c in ["Mecânico (Nome)", "Nome", "Técnico", "Tecnico", "Operário", "Operario"] if c in df_ausencias.columns), df_ausencias.columns[3] if len(df_ausencias.columns) > 3 else df_ausencias.select_dtypes(include=['object']).columns[0])
        col_val_aus = df_ausencias.select_dtypes(include=['number']).columns[-1]

        if col_desc_aus and col_val_aus:
            df_aus_filtrado = df_ausencias[df_ausencias[col_desc_aus].astype(str).str.strip().str.lower() != "presença"]
            df_aus_filtrado["Técnico_clean"] = df_aus_filtrado[col_op_aus].astype(str).str.strip().str.upper()
            df_ausencias_grouped = df_aus_filtrado.groupby("Técnico_clean", as_index=False)[col_val_aus].sum()
            df_ausencias_grouped.rename(columns={col_val_aus: "Ausencias_calc"}, inplace=True)
        else:
            df_ausencias_grouped = pd.DataFrame(columns=["Técnico_clean", "Ausencias_calc"])
    else:
        df_ausencias_grouped = pd.DataFrame(columns=["Técnico_clean", "Ausencias_calc"])
