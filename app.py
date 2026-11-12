import pandas as pd
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="Relatório de Produtividade - Oficinas", layout="wide")

st.title("Relatório de Produtividade de Oficinas")

meses_pt = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto", 
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}

# 1. SECÇÃO: PARÂMETROS DO MÊS
st.sidebar.header("Parâmetros do Mês")
mes_atual_num = datetime.now().month

mes_selecionado = st.sidebar.selectbox(
    "Mês Corrente:", 
    options=list(meses_pt.keys()), 
    format_func=lambda x: meses_pt[x], 
    index=list(meses_pt.keys()).index(mes_atual_num) if mes_atual_num in meses_pt else 10
)
ano_corrente = 2026

if st.sidebar.button("Atualizar / Limpar Cache"):
    st.cache_data.clear()
    st.rerun()

dias_uteis_iniciais = 22
decorridos_iniciais = 22 if (mes_selecionado < mes_atual_num) else min(8, 22)

dias_uteis_padrao_mes = st.sidebar.slider(
    f"Total de Dias Úteis ({meses_pt[mes_selecionado]}):", 
    min_value=18, max_value=25, value=dias_uteis_iniciais
)

dias_decorridos = st.sidebar.number_input(
    "Dias úteis decorridos:", 
    min_value=1, max_value=dias_uteis_padrao_mes, 
    value=dias_uteis_padrao_mes if (mes_selecionado < mes_atual_num) else decorridos_iniciais
)

percentagem_progresso = (dias_decorridos / dias_uteis_padrao_mes) * 100 if dias_uteis_padrao_mes > 0 else 0
progresso_decimal = min(max(dias_decorridos / dias_uteis_padrao_mes, 0.0), 1.0)

st.sidebar.progress(
    progresso_decimal,
    text=f"Progresso: {dias_decorridos} de {dias_uteis_padrao_mes} dias ({percentagem_progresso:.1f}%)"
)

# 2. RESERVAR ESPAÇO PARA OS FILTROS (SEMPRE ACIMA DOS UPLOADS)
st.sidebar.markdown("---")
container_filtros = st.sidebar.container()

# 3. SECÇÃO: CARREGAR FICHEIROS EXCEL (NO FUNDO)
st.sidebar.markdown("---")
st.sidebar.header("Carregar Ficheiros Excel")
file_mecanicos = st.sidebar.file_uploader("Carregar 'tabela mecânicos.xlsx'", type=["xlsx"])
file_hf = st.sidebar.file_uploader("Carregar 'Comparativo de Horas'", type=["xlsx"])
file_ausencias = st.sidebar.file_uploader("Carregar 'Listagem de Ausencias'", type=["xlsx"])

if file_mecanicos and file_hf:
    xl = pd.ExcelFile(file_mecanicos)
    df_mecanicos = xl.parse("Mecânicos" if "Mecânicos" in xl.sheet_names else xl.sheet_names[0])
    df_hf = pd.read_excel(file_hf)
    df_ausencias = pd.read_excel(file_ausencias) if file_ausencias else pd.DataFrame()

    df_mecanicos.rename(columns={"Técnicos": "Técnico", "Nome": "Técnico"}, inplace=True)
    df_mecanicos["Técnico_clean"] = df_mecanicos["Técnico"].astype(str).str.strip().str.upper()
    df_mecanicos["Oficina"] = df_mecanicos["Oficina"].astype(str).str.strip()
    df_mecanicos["Local"] = df_mecanicos["Local"].astype(str).str.strip()

    oficinas_oficiais = df_mecanicos["Oficina"].unique().tolist()
    dict_oficina_local = dict(zip(df_mecanicos["Oficina"].str.upper(), df_mecanicos["Local"]))

    def normalizar_oficina(nome_oficina):
        if pd.isna(nome_oficina):
            return "Geral"
        nome_str = str(nome_oficina).strip()
        nome_upper = nome_str.upper()
        for oficina_oficial in oficinas_oficiais:
            if oficina_oficial.upper() == nome_upper:
                return oficina_oficial
        if "RENAULT BEJA" in nome_upper or "PEUGEOT BEJA" in nome_upper or "MECÂNICA BEJA" in nome_upper:
            return "MECÂNICA BEJA"
        if "COLISÃO BEJA" in nome_upper or "COLISAO BEJA" in nome_upper:
            return "COLISÃO BEJA"
        return nome_str

    col_operario_hf = next((c for c in ["Operário", "Operario", "Técnico", "Tecnico", "Nome", "Mecânico"] if c in df_hf.columns), df_hf.select_dtypes(include=['object']).columns[0])
    col_fact = next((c for c in ["Tempos facturados", "Tempos facturados ", "Horas facturadas", "Facturado", "Valor"] if c in df_hf.columns), df_hf.select_dtypes(include=['number']).columns[-1])
    col_oficina_hf = next((c for c in ["Oficina", "Escritório", "Escritorio", "Local"] if c in df_hf.columns), "Oficina")

    df_hf["Operario_clean"] = df_hf[col_operario_hf].astype(str).str.strip().str.upper()
    df_hf["Oficina_Normalizada"] = df_hf[col_oficina_hf].apply(normalizar_oficina)

    df_hf_known = df_hf[df_hf["Operario_clean"].notna() & (df_hf["Operario_clean"] != "") & (df_hf["Operario_clean"] != "NAN")]
    df_hf_grouped = df_hf_known.groupby("Operario_clean", as_index=False)[col_fact].sum()
    df_hf_grouped.rename(columns={col_fact: "Horas facturadas", "Operario_clean": "Técnico_clean"}, inplace=True)

    df_merged = pd.merge(df_mecanicos, df_hf_grouped, on="Técnico_clean", how="left")
    df_merged["Horas facturadas"] = df_merged["Horas facturadas"].fillna(0.0)
    df_merged["Ausencias"] = df_merged["Ausencias"].fillna(0.0) if "Ausencias" in df_merged.columns else 0.0
    df_merged["Horas reais dia"] = df_merged["Horas reais dia"].fillna(7.5)
    df_merged["Horas disponiveis"] = (dias_decorridos * df_merged["Horas reais dia"]) - df_merged["Ausencias"]

    colunas_finais = [
        "Técnico", "Local", "Oficina", "Função", 
        "Horas potencial dia", "Horas reais dia", 
        "Horas disponiveis", "Horas facturadas", "Produtividade %"
    ]
    
    df_merged["Produtividade %"] = 0.0
    mask_disp = df_merged["Horas disponiveis"] > 0
    df_merged.loc[mask_disp, "Produtividade %"] = (
        df_merged.loc[mask_disp, "Horas facturadas"] / df_merged.loc[mask_disp, "Horas disponiveis"]
    ) * 100

    df_exibicao_tecnicos = df_merged[[c for c in colunas_finais if c in df_merged.columns]].copy()

    df_hf_sem_op = df_hf[df_hf["Operario_clean"].isna() | (df_hf["Operario_clean"] == "") | (df_hf["Operario_clean"] == "NAN")]
    
    linhas_sem_produtivo = []
    if not df_hf_sem_op.empty:
        df_sem_op_grouped = df_hf_sem_op.groupby("Oficina_Normalizada", as_index=False)[col_fact].sum()
        for _, row in df_sem_op_grouped.iterrows():
            oficina_val = row["Oficina_Normalizada"]
            horas_fact_val = row[col_fact]
            
            if horas_fact_val > 0:
                local_val = dict_oficina_local.get(oficina_val.upper(), "Outro")
                if local_val == "Outro":
                    of_upper = oficina_val.upper()
                    if "BEJA" in of_upper:
                        local_val = "Beja"
                    elif "ÉVORA" in of_upper or "EVORA" in of_upper:
                        local_val = "Évora"
                    elif "FARO" in of_upper:
                        local_val = "Faro"
                    elif "SETÚBAL" in of_upper or "SETUBAL" in of_upper:
                        local_val = "Setúbal"
                    elif "SINES" in of_upper:
                        local_val = "Sines"
                    elif "PORTIMÃO" in of_upper or "PORTIMAO" in of_upper:
                        local_val = "Portimão"

                linhas_sem_produtivo.append({
                    "Técnico": "Sem produtivo",
                    "Local": local_val,
                    "Oficina": oficina_val,
                    "Função": "Sem Produtivo",
                    "Horas potencial dia": 0.0,
                    "Horas reais dia": 0.0,
                    "Horas disponiveis": 0.0,
                    "Horas facturadas": horas_fact_val,
                    "Produtividade %": 0.0
                })

    df_sem_prod_final = pd.DataFrame(linhas_sem_produtivo)
    df_completo = pd.concat([df_exibicao_tecnicos, df_sem_prod_final], ignore_index=True) if not df_sem_prod_final.empty else df_exibicao_tecnicos

    with container_filtros:
        st.header("Filtros")
        locais_disponiveis = ["Todas"] + sorted(df_completo["Local"].dropna().astype(str).unique().tolist())
        sel_local = st.selectbox("Selecionar Cidade (Local)", locais_disponiveis)

        df_filtrado_local = df_completo[df_completo["Local"] == sel_local] if sel_local != "Todas" else df_completo

        oficinas_disponiveis = ["Todas"] + sorted(df_filtrado_local["Oficina"].dropna().astype(str).unique().tolist())
        sel_oficina = st.selectbox("Selecionar Oficina", oficinas_disponiveis)

    df_final_filtro = df_filtrado_local[df_filtrado_local["Oficina"] == sel_oficina] if sel_oficina != "Todas" else df_filtrado_local

    tot_h_potencial = df_final_filtro["Horas potencial dia"].sum() if "Horas potencial dia" in df_final_filtro.columns else 0.0
    tot_h_reais = df_final_filtro["Horas reais dia"].sum() if "Horas reais dia" in df_final_filtro.columns else 0.0
    tot_h_disp = df_final_filtro["Horas disponiveis"].sum()
    tot_h_fact = df_final_filtro["Horas facturadas"].sum()
    tot_prod = (tot_h_fact / tot_h_disp * 100) if tot_h_disp > 0 else 0.0

    linha_total = {
        "Técnico": "TOTAL",
        "Local": "-",
        "Oficina": "-",
        "Função": "-",
        "Horas potencial dia": tot_h_potencial,
        "Horas reais dia": tot_h_reais,
        "Horas disponiveis": tot_h_disp,
        "Horas facturadas": tot_h_fact,
        "Produtividade %": tot_prod
    }

    df_exibicao_final = pd.concat([df_final_filtro, pd.DataFrame([linha_total])], ignore_index=True)

    def destacar_total(s):
        is_total = s.index.isin(df_exibicao_final[df_exibicao_final["Técnico"] == "TOTAL"].index)
        return ['font-weight: bold; background-color: #f0f2f6;' if v else '' for v in is_total]

    st.subheader(f"Mapa de Produtividade - {meses_pt[mes_selecionado]} {ano_corrente}")
    st.dataframe(
        df_exibicao_final.style.apply(destacar_total, axis=1).format({
            "Horas potencial dia": "{:.1f}",
            "Horas reais dia": "{:.1f}",
            "Horas disponiveis": "{:.1f}",
            "Horas facturadas": "{:.2f}",
            "Produtividade %": "{:.1f}%"
        }, na_rep="-"),
        use_container_width=True
    )
else:
    with container_filtros:
        st.header("Filtros")
        st.caption("Carregue os ficheiros abaixo para ativar os filtros.")
    st.info("👈 Por favor, carregue a **tabela de mecânicos** e o **comparativo de horas** no fundo da barra lateral para gerar o relatório.")
