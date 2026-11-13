import pandas as pd
import streamlit as st
from datetime import datetime, date, timedelta
import os
import glob
import altair as alt
import unicodedata
import io

st.set_page_config(page_title="Relatório de Produtividade e Faturação - Oficinas", layout="wide")

st.markdown("""
    <style>
    div.stButton > button:first-child {
        background-color: #9ca3af;
        color: #ffffff;
        border: 1px solid #4b5563;
        font-weight: 600;
    }
    div.stButton > button:first-child:hover {
        background-color: #6b7280;
        color: #ffffff;
        border: 1px solid #374151;
    }
    </style>
""", unsafe_allow_html=True)

col_titulo, col_logo = st.columns([4, 1])
with col_titulo:
    st.title("Relatório de Oficinas - Grupo Motorpor")
with col_logo:
    logo_path = "logo_motorpor.jpg"
    if os.path.exists(logo_path):
        st.image(logo_path, width=150)

meses_pt = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 
    5: "maio", 6: "junho", 7: "julho", 8: "agosto", 
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
}

meses_sigla = {
    1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 
    5: "MAI", 6: "JUN", 7: "JUL", 8: "AGO", 
    9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"
}

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.utilizador = ""
    st.session_state.oficinas_autorizadas = []
    st.session_state.acesso_total = False

df_perfis = pd.DataFrame()
if os.path.exists("perfis_utilizadores.xlsx"):
    try:
        df_perfis = pd.read_excel("perfis_utilizadores.xlsx")
        df_perfis["Utilizador_clean"] = df_perfis["Utilizador"].astype(str).str.strip().str.lower()
        df_perfis["password_clean"] = df_perfis["password"].astype(str).str.strip()
    except Exception as e:
        st.error(f"Erro ao ler perfis_utilizadores.xlsx: {e}")

st.sidebar.header("Navegação e Módulo")
modulo_selecionado = st.sidebar.radio(
    "Escolha o Módulo:",
    options=["📊 Produtividade de Oficinas", "💶 Faturação e Orçamento Diário"]
)

st.sidebar.markdown("---")
st.sidebar.header("Seleção de Período")
ano_corrente = st.sidebar.selectbox("Ano:", options=[2025, 2026, 2027], index=1)
mes_atual_num = datetime.now().month

mes_selecionado = st.sidebar.selectbox(
    "Mês:", 
    options=list(meses_pt.keys()), 
    format_func=lambda x: meses_pt[x].capitalize(), 
    index=list(meses_pt.keys()).index(mes_atual_num) if (mes_atual_num in meses_pt and ano_corrente == 2026) else 10
)

sigla_mes = meses_sigla.get(mes_selecionado, "NOV")

cidade_feriado = st.sidebar.selectbox(
    "Local para Feriados Municipais:",
    options=["Nacional (Sem Municipal)", "Beja", "Évora", "Portimão", "Faro", "Setúbal", "Sines"]
)

if st.sidebar.button("Atualizar / Limpar Cache"):
    st.cache_data.clear()
    st.rerun()

ultima_atualizacao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
st.sidebar.caption(f"🔄 Última atualização: {ultima_atualizacao}")

def calcular_dias_uteis_mes(ano, mes, cidade="Nacional (Sem Municipal)"):
    if mes == 12:
        proximo_mes = date(ano + 1, 1, 1)
    else:
        proximo_mes = date(ano, mes + 1, 1)
    inicio = date(ano, mes, 1)
    fim = proximo_mes - timedelta(days=1)
    
    feriados = [
        date(ano, 1, 1), date(ano, 4, 25), date(ano, 5, 1),   
        date(ano, 6, 10), date(ano, 8, 15), date(ano, 10, 5),  
        date(ano, 12, 1), date(ano, 12, 8), date(ano, 12, 25), 
    ]
    if ano == 2026:
        feriados.extend([date(2026, 4, 3), date(2026, 6, 4)])
        
    cidade_upper = str(cidade).upper()
    if "BEJA" in cidade_upper: feriados.append(date(ano, 5, 14))    
    elif "ÉVORA" in cidade_upper or "EVORA" in cidade_upper: feriados.append(date(ano, 6, 29))    
    elif "PORTIMÃO" in cidade_upper or "PORTIMAO" in cidade_upper: feriados.append(date(ano, 12, 19))   
    elif "FARO" in cidade_upper: feriados.append(date(ano, 9, 7))     
    elif "SETÚBAL" in cidade_upper or "SETUBAL" in cidade_upper: feriados.append(date(ano, 9, 15))    
    elif "SINES" in cidade_upper: feriados.append(date(ano, 11, 24))   
        
    dias_uteis = 0
    atual = inicio
    while atual <= fim:
        if atual.weekday() < 5 and atual not in feriados:
            dias_uteis += 1
        atual += timedelta(days=1)
    return dias_uteis

dias_uteis_calculados = calcular_dias_uteis_mes(ano_corrente, mes_selecionado, cidade_feriado)
dias_uteis_padrao_mes = st.sidebar.slider(f"Total de Dias Úteis:", min_value=15, max_value=25, value=dias_uteis_calculados)
dias_decorridos = st.sidebar.number_input("Dias úteis decorridos:", min_value=1, max_value=dias_uteis_padrao_mes, value=8)

st.sidebar.markdown("---")
container_filtros = st.sidebar.container()

utilizadores_com_permissao_upload = ["vazul", "mchegadinho"]
utilizador_atual = st.session_state.utilizador.strip().lower() if st.session_state.autenticado else ""
tem_permissao_upload = st.session_state.autenticado and (utilizador_atual in [u.lower() for u in utilizadores_com_permissao_upload])

if tem_permissao_upload:
    st.sidebar.markdown("---")
    st.sidebar.header("Atualizar Ficheiros (Admin)")
    st.sidebar.file_uploader("Substituir 'tabela mecânicos'", type=["xlsx"], key="up_mec")
    st.sidebar.file_uploader("Substituir 'Comparativo de Horas'", type=["xlsx"], key="up_hf")
    st.sidebar.file_uploader("Substituir 'Listagem de Ausencias'", type=["xlsx"], key="up_aus")
    st.sidebar.file_uploader("Mapa de Faturação 1 (CSV)", type=["csv"], key="up_f1")
    st.sidebar.file_uploader("Mapa de Faturação 2 (CSV)", type=["csv"], key="up_f2")
    st.sidebar.file_uploader("Orçamento e Agrupamentos (XLSX)", type=["xlsx"], key="up_orc")

st.sidebar.markdown("---")
container_login = st.sidebar.container()
with container_login:
    st.header("Acesso / Login")
    if not st.session_state.autenticado:
        if not df_perfis.empty:
            user_input = st.text_input("Utilizador:").strip().lower()
            pass_input = st.text_input("Password:", type="password").strip()
            if st.button("Entrar"):
                if user_input and pass_input:
                    match = df_perfis[(df_perfis["Utilizador_clean"] == user_input) & (df_perfis["password_clean"] == pass_input)]
                    if not match.empty:
                        st.session_state.autenticado = True
                        st.session_state.utilizador = match.iloc[0]["Utilizador"]
                        permissoes_str = str(match.iloc[0]["Oficinas permitidas"]).strip()
                        st.session_state.acesso_total = (permissoes_str.lower() == "todas")
                        st.session_state.oficinas_autorizadas = [] if st.session_state.acesso_total else [o.strip().upper() for o in permissoes_str.split(",") if o.strip()]
                        st.rerun()
                    else:
                        st.error("Utilizador ou password incorretos.")
        else:
            st.error("Ficheiro 'perfis_utilizadores.xlsx' não encontrado.")
    else:
        st.success(f"Sessão iniciada como: **{st.session_state.utilizador}**")
        if st.button("Logout"):
            st.session_state.autenticado = False
            st.rerun()

file_mecanicos_up = st.session_state.get("up_mec")
file_hf_up = st.session_state.get("up_hf")
file_fact1_up = st.session_state.get("up_f1")
file_fact2_up = st.session_state.get("up_f2")
file_orc_up = st.session_state.get("up_orc")

path_mecanicos = file_mecanicos_up if file_mecanicos_up is not None else ("tabela mecânicos.xlsx" if os.path.exists("tabela mecânicos.xlsx") else None)
path_hf = file_hf_up if file_hf_up is not None else (f"Comparativo de Horas_{sigla_mes}_{ano_corrente}.xlsx" if os.path.exists(f"Comparativo de Horas_{sigla_mes}_{ano_corrente}.xlsx") else None)

if not st.session_state.autenticado:
    st.warning("🔒 Por favor, faça login na barra lateral.")
elif not path_mecanicos or not path_hf:
    st.warning("⚠️ Ficheiros padrão em falta.")
else:
    if modulo_selecionado == "📊 Produtividade de Oficinas":
        st.info("Módulo de Produtividade ativo.")
    elif modulo_selecionado == "💶 Faturação e Orçamento Diário":
        st.subheader(f"Faturação e Orçamento Diário - {meses_pt[mes_selecionado].capitalize()} {ano_corrente}")
        
        excel_orc_path = "Orçamento-agrupamento oficinas.xlsx"
        df_orc, df_agrup = pd.DataFrame(), pd.DataFrame()
        target_orc = file_orc_up if file_orc_up is not None else (excel_orc_path if os.path.exists(excel_orc_path) else None)
        
        if target_orc is not None:
            try:
                df_orc = pd.read_excel(target_orc, sheet_name="Oficina_ORÇAMENTO")
                df_agrup = pd.read_excel(target_orc, sheet_name="Agrupamento_oficinas")
            except Exception as e:
                st.error(f"Erro ao ler orçamentos: {e}")
        
        df_fact_total = pd.DataFrame()
        csv_fontes = [file_fact1_up, file_fact2_up]
        if all(f is None for f in csv_fontes):
            csv_fontes = glob.glob("*.csv")
            
        for file_f in csv_fontes:
            if file_f is not None:
                try:
                    file_bytes = io.BytesIO(file_f.getvalue()) if hasattr(file_f, "getvalue") else open(file_f, "rb")
                    temp_df = pd.read_csv(file_bytes, sep=None, engine="python", encoding="iso-8859-1", on_bad_lines='skip', header=0)
                    if not temp_df.empty:
                        df_fact_total = pd.concat([df_fact_total, temp_df], ignore_index=True)
                except Exception:
                    pass
        
        if df_fact_total.empty:
            st.warning("⚠️ Nenhum ficheiro CSV de faturação encontrado ou carregado.")
        else:
            with st.container():
                st.success("✅ Ficheiros CSV lidos com sucesso! Veja abaixo o diagnóstico das colunas:")
                st.write("Colunas detetadas no CSV:", list(df_fact_total.columns))
                st.dataframe(df_fact_total.head(2))

            idx_ofic = 0
            df_fact_total["Oficina_CSV"] = df_fact_total.iloc[:, idx_ofic].astype(str).str.strip()
            
            oficinas_disponiveis_fat = sorted([o for o in df_fact_total["Oficina_CSV"].unique() if o and "MAPA" not in o.upper() and "OFICINA" not in o.upper()])
            if not oficinas_disponiveis_fat:
                oficinas_disponiveis_fat = ["Geral"]

            oficina_selecionada_fat = st.selectbox("Selecionar Oficina (Faturação):", options=oficinas_disponiveis_fat)
            
            col_graf, col_kpi = st.columns([3, 1])
            with col_kpi:
                st.metric(label="ORÇAMENTO MENSAL", value="0 €")
                st.metric(label="FATURAÇÃO À DATA", value="0 €")
