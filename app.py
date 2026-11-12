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

    df_hf_known = df_hf[df_hf["Operario_clean"].notna() & (df_hf["Operario_clean"]
