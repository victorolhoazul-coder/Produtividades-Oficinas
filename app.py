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

# Inicializar o estado de autenticação
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.utilizador = ""
    st.session_state.oficinas_autorizadas = []
    st.session_state.acesso_total = False

# 0. RESERVAR ESPAÇO PARA O LOGIN / LOGOUT NO TOPO DA BARRA LATERAL
container_login = st.sidebar.container()

# 1. SECÇÃO: PARÂMETROS DO MÊS
st.sidebar.markdown("---")
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

# 2. RESERVAR ESPAÇO PARA OS FILTROS
st.sidebar.markdown("---")
container_filtros = st.sidebar.container()

# 3. SECÇÃO: CARREGAR FICHEIROS EXCEL (NO FUNDO)
st.sidebar.markdown("---")
st.sidebar.header("Carregar Ficheiros Excel")
file_perfis = st.sidebar.file_uploader("Carregar 'perfis_utilizadores.xlsx'", type=["xlsx"])
file_mecanicos = st.sidebar.file_uploader("Carregar 'tabela mecânicos.xlsx'", type=["xlsx"])
file_hf = st.sidebar.file_uploader("Carregar 'Comparativo de Horas'", type=["xlsx"])
file_ausencias = st.sidebar.file_uploader("Carregar 'Listagem de Ausencias'", type=["xlsx"])

# LÓGICA DE AUTENTICAÇÃO E LOGOUT
with container_login:
    st.header("Acesso / Login")
    if file_perfis:
        df_perfis = pd.read_excel(file_perfis)
        df_perfis["Utilizador_clean"] = df_perfis["Utilizador"].astype(str).str.strip().str.lower()
        df_perfis["password_clean"] = df_perfis["password"].astype(str).str.strip()

        if not st.session_state.autenticado:
            user_input = st.text_input("Utilizador:").strip().lower()
            pass_input = st.text_input("Password:", type="password").strip()

            if st.button("Entrar"):
                if user_input and pass_input:
                    match = df_perfis[(df_perfis["Utilizador_clean"] == user_input) & (df_perfis["password_clean"] == pass_input)]
                    if not match.empty:
                        st.session_state.autenticado = True
                        st.session_state.utilizador = match.iloc[0]["Utilizador"]
                        permissoes_str = str(match.iloc[0]["Oficinas permitidas"]).strip()
                        
                        if permissoes_str.lower() == "todas":
                            st.session_state.acesso_total = True
                            st.session_state.oficinas_autorizadas = []
                        else:
                            st.session_state.acesso_total = False
                            st.session_state.oficinas_autorizadas = [o.strip().upper() for o in permissoes_str.split(",") if o.strip()]
                        st.rerun()
                    else:
                        st.error("Utilizador ou password incorretos.")
                else:
                    st.warning("Preencha o utilizador e a password.")
        else:
            st.success(f"Sessão iniciada como: **{st.session_state.utilizador}**")
            if st.button("Logout / Mudar Utilizador"):
                st.session_state.autenticado = False
                st.session_state.utilizador = ""
                st.session_state.oficinas_autorizadas = []
                st.session_state.acesso_total = False
                st.rerun()
    else:
        st.caption("Carrege o ficheiro de perfis no fundo da barra lateral.")

# PROCESSAMENTO DOS DADOS
if not file_perfis:
    st.info("👈 Por favor, carregue o ficheiro **perfis_utilizadores.xlsx** no fundo da barra lateral.")
elif not st.session_state.autenticado:
    st.warning("🔒 Por favor, faça login no topo da barra lateral para aceder aos relatórios.")
elif file_mecanicos and file_hf:
    xl = pd.ExcelFile(file_mecanicos)
    df_mecanicos = xl.parse("Mecânicos" if "Mecânicos" in xl.sheet_names else xl.sheet_names[0])
    df_hf = pd.read_excel(file_hf)
    df_ausencias = pd.read_excel(file_ausencias) if file_ausencias else pd.DataFrame()

    df_mecanicos.rename(columns={"Técnicos": "Técnico", "Nome": "Técnico"}, inplace=True)
    df_mecanicos["Técnico_clean"] = df_mecanicos["T
