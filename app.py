import pandas as pd
import streamlit as st
from datetime import datetime
import os

st.set_page_config(page_title="Relatório de Produtividade - Oficinas", layout="wide")

st.title("Relatório de Produtividade de Oficinas")

meses_pt = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto", 
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}

meses_sigla = {
    1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 
    5: "MAI", 6: "JUN", 7: "JUL", 8: "AGO", 
    9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"
}

# Inicializar o estado de autenticação
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.utilizador = ""
    st.session_state.oficinas_autorizadas = []
    st.session_state.acesso_total = False

# CARREGAR PERFIS AUTOMATICAMENTE DO REPOSITÓRIO
df_perfis = pd.DataFrame()
if os.path.exists("perfis_utilizadores.xlsx"):
    try:
        df_perfis = pd.read_excel("perfis_utilizadores.xlsx")
        df_perfis["Utilizador_clean"] = df_perfis["Utilizador"].astype(str).str.strip().str.lower()
        df_perfis["password_clean"] = df_perfis["password"].astype(str).str.strip()
    except Exception as e:
        st.error(f"Erro ao ler perfis_utilizadores.xlsx: {e}")

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

# 3. SECÇÃO: CARREGAR FICHEIROS (APENAS PARA ADMINS ATUALIZAREM)
utilizadores_com_permissao_upload = ["vazul", "mchegadinho"]
utilizador_atual = st.session_state.utilizador.strip().lower() if st.session_state.autenticado else ""
tem_permissao_upload = st.session_state.autenticado and (utilizador_atual in [u.lower() for u in utilizadores_com_permissao_upload])

file_mecanicos_up = None
file_hf_up = None
file_ausencias_up = None

if tem_permissao_upload:
    st.sidebar.markdown("---")
    st.sidebar.header("Atualizar Ficheiros (Admin)")
    file_mecanicos_up = st.sidebar.file_uploader("Substituir 'tabela mecânicos'", type=["xlsx"])
    file_hf_up = st.sidebar.file_uploader("Substituir 'Comparativo de Horas'", type=["xlsx"])
    file_ausencias_up = st.sidebar.file_uploader("Substituir 'Listagem de Ausencias'", type=["xlsx"])

# LÓGICA DE AUTENTICAÇÃO E LOGOUT
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
            st.error("Ficheiro 'perfis_utilizadores.xlsx' não encontrado no repositório GitHub.")
    else:
        st.success(f"Sessão iniciada como: **{st.session_state.utilizador}**")
        if st.button("Logout / Mudar Utilizador"):
            st.session_state.autenticado = False
            st.session_state.utilizador = ""
            st.session_state.oficinas_autorizadas = []
            st.session_state.acesso_total = False
            st.rerun()

# PROCESSAMENTO DOS DADOS
sigla_mes = meses_sigla.get(mes_selecionado, "NOV")
path_mecanicos = file_mecanicos_up if file_mecanicos_up else ("tabela mecânicos.xlsx" if os.path.exists("tabela mecânicos.xlsx") else None)
path_hf = file_hf_up if file_hf_up else (f"Comparativo de Horas_{sigla_mes}_{ano_corrente}.xlsx" if os.path.exists(f"Comparativo de Horas_{sigla_mes}_{ano_corrente}.xlsx") else None)
path_ausencias = f"Listagem de AusenciasPresenças_{sigla_mes}_{ano_corrente}.xlsx" if os.path.exists(f"Listagem de AusenciasPresenças_{sigla_mes}_{ano_corrente}.xlsx") else None

if not st.session_state.autenticado:
    st.warning("🔒 Por favor, faça login no topo da barra lateral para aceder aos relatórios.")
elif not path_mecanicos or not path_hf:
    with container_filtros:
        st.header("Filtros")
        st.caption("Aguardando ficheiros de dados.")
    st.warning(f"⚠️ Não foram encontrados os ficheiros padrão no repositório para o mês de {meses_pt[mes_selecionado]} ({sigla_mes}). Certifique-se de que os carregou para o GitHub ou use o painel de admin.")
else:
    xl = pd.ExcelFile(path_mecanicos)
    df_mecanicos = xl.parse("Mecânicos" if "Mecânicos" in xl.sheet_names else xl.sheet_names[0])
    df_hf = pd.read_excel(path_hf)
    
    # PROCESSAMENTO DE AUSÊNCIAS (COM CORREÇÃO ROBUSTA)
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

    df_mecanicos.rename(columns={"Técnicos": "Técnico", "Nome": "Técnico"}, inplace=True)
    df_mecanicos["Técnico_clean"] = df_mecanicos["Técnico"].astype(str).str.strip().str.upper()
    df_mecanicos["Oficina"] = df_mecanicos["Oficina"].astype(str).str.strip()
    df_mecanicos["Local"] = df_mecanicos["Local"].astype(str).str.strip()

    oficinas_oficiais = df_mecanicos["Oficina"].unique().tolist()
    dict_oficina_local = dict(zip(df_mecanicos["Oficina"].str.upper(), df_mecanicos["Local"]))

    oficina_portimao_oficial = "MECÂNICA PORTIMÃO"
    for o in oficinas_oficiais:
        if "PORTIMÃO" in o.upper() or "PORTIMAO" in o.upper():
            oficina_portimao_oficial = o
            break

    def normalizar_oficina(nome_oficina):
        if pd.isna(nome_oficina):
            return "Geral"
        nome_str = str(nome_oficina).strip()
        nome_upper = nome_str.upper()
        
        for oficina_oficial in oficinas_oficiais:
            if oficina_oficial.upper() == nome_upper:
                return oficina_oficial
                
        if "BEJA" in nome_upper:
            return "MECÂNICA BEJA"
        if "PORTIMÃO" in nome_upper or "PORTIMAO" in nome_upper:
            return oficina_portimao_oficial
        if "FARO" in nome_upper:
            for o in oficinas_oficiais:
                if "FARO" in o.upper(): return o
        if "ÉV
