import pandas as pd
import streamlit as st
from datetime import datetime, date, timedelta
import os
import glob
import altair as alt

st.set_page_config(page_title="Relatório de Produtividade e Faturação - Oficinas", layout="wide")

# CSS personalizado para escurecer o botão de atualizar/limpar cache
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

# Cabeçalho com Título e Logotipo no canto superior direito
col_titulo, col_logo = st.columns([4, 1])
with col_titulo:
    st.title("Relatório de Oficinas - Grupo Motorpor")
with col_logo:
    logo_path = "logo_motorpor.jpg"
    if os.path.exists(logo_path):
        st.image(logo_path, width=150)

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

container_login = st.sidebar.container()

st.sidebar.markdown("---")
st.sidebar.header("Seleção de Período")
ano_corrente = st.sidebar.selectbox("Ano:", options=[2025, 2026, 2027], index=1)
mes_atual_num = datetime.now().month

mes_selecionado = st.sidebar.selectbox(
    "Mês:", 
    options=list(meses_pt.keys()), 
    format_func=lambda x: meses_pt[x], 
    index=list(meses_pt.keys()).index(mes_atual_num) if (mes_atual_num in meses_pt and ano_corrente == 2026) else 10
)

sigla_mes = meses_sigla.get(mes_selecionado, "NOV")

# Seletor para o feriado municipal por cidade
cidade_feriado = st.sidebar.selectbox(
    "Local para Feriados Municipais:",
    options=["Nacional (Sem Municipal)", "Beja", "Évora", "Portimão", "Faro", "Setúbal", "Sines"]
)

if st.sidebar.button("Atualizar / Limpar Cache"):
    st.cache_data.clear()
    st.rerun()

ultima_atualizacao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
st.sidebar.caption(f"🔄 Última atualização: {ultima_atualizacao}")

# Função para calcular dias úteis reais do mês selecionado (Nacionais + Municipais)
def calcular_dias_uteis_mes(ano, mes, cidade="Nacional (Sem Municipal)"):
    if mes == 12:
        proximo_mes = date(ano + 1, 1, 1)
    else:
        proximo_mes = date(ano, mes + 1, 1)
    inicio = date(ano, mes, 1)
    fim = proximo_mes - timedelta(days=1)
    
    feriados = [
        date(ano, 1, 1),   # Ano Novo
        date(ano, 4, 25),  # Dia da Liberdade
        date(ano, 5, 1),   # Dia do Trabalhador
        date(ano, 6, 10),  # Dia de Portugal
        date(ano, 8, 15),  # Assunção de Nossa Senhora
        date(ano, 10, 5),  # Implantação da República
        date(ano, 12, 1),  # Restauração da Independência
        date(ano, 12, 8),  # Imaculada Conceição
        date(ano, 12, 25), # Natal
    ]
    
    if ano == 2026:
        feriados.extend([
            date(2026, 4, 3),   # Sexta-Feira Santa
            date(2026, 6, 4),   # Corpo de Deus
        ])
        
    cidade_upper = str(cidade).upper()
    if "BEJA" in cidade_upper:
        feriados.append(date(ano, 5, 14))    # Quinta-feira da Ascensão
    elif "ÉVORA" in cidade_upper or "EVORA" in cidade_upper:
        feriados.append(date(ano, 6, 29))    # São Pedro
    elif "PORTIMÃO" in cidade_upper or "PORTIMAO" in cidade_upper:
        feriados.append(date(ano, 12, 19))   # Dia de Portimão
    elif "FARO" in cidade_upper:
        feriados.append(date(ano, 9, 7))     # Dia de Faro
    elif "SETÚBAL" in cidade_upper or "SETUBAL" in cidade_upper:
        feriados.append(date(ano, 9, 15))    # Dia de Setúbal
    elif "SINES" in cidade_upper:
        feriados.append(date(ano, 11, 24))   # Dia de Sines
        
    dias_uteis = 0
    atual = inicio
    while atual <= fim:
        if atual.weekday() < 5 and atual not in feriados:
            dias_uteis += 1
        atual += timedelta(days=1)
    return dias_uteis

dias_uteis_calculados = calcular_dias_uteis_mes(ano_corrente, mes_selecionado, cidade_feriado)
decorridos_iniciais = dias_uteis_calculados if (mes_selecionado < mes_atual_num or ano_corrente < 2026) else min(8, dias_uteis_calculados)

dias_uteis_padrao_mes = st.sidebar.slider(
    f"Total de Dias Úteis ({meses_pt[mes_selecionado]}):", 
    min_value=15, max_value=25, value=dias_uteis_calculados
)

dias_decorridos = st.sidebar.number_input(
    "Dias úteis decorridos:", 
    min_value=1, max_value=dias_uteis_padrao_mes, 
    value=dias_uteis_padrao_mes if (mes_selecionado < mes_atual_num or ano_corrente < 2026) else decorridos_iniciais
)

percentagem_progresso = (dias_decorridos / dias_uteis_padrao_mes) * 100 if dias_uteis_padrao_mes > 0 else 0
progresso_decimal = min(max(dias_decorridos / dias_uteis_padrao_mes, 0.0), 1.0)

st.sidebar.progress(
    progresso_decimal,
    text=f"Progresso: {dias_decorridos} de {dias_uteis_padrao_mes} dias ({percentagem_progresso:.1f}%)"
)

st.sidebar.markdown("---")
container_filtros = st.sidebar.container()

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

path_mecanicos = file_mecanicos_up if file_mecanicos_up else ("tabela mecânicos.xlsx" if os.path.exists("tabela mecânicos.xlsx") else None)
path_hf = file_hf_up if file_hf_up else (f"Comparativo de Horas_{sigla_mes}_{ano_corrente}.xlsx" if os.path.exists(f"Comparativo de Horas_{sigla_mes}_{ano_corrente}.xlsx") else None)
path_ausencias = f"Listagem de AusenciasPresenças_{sigla_mes}_{ano_corrente}.xlsx" if os.path.exists(f"Listagem de AusenciasPresenças_{sigla_mes}_{ano_corrente}.xlsx") else None

if not st.session_state.autenticado:
    st.warning("🔒 Por favor, faça login no topo da barra lateral para aceder aos relatórios.")
elif not path_mecanicos or not path_hf:
    with container_filtros:
        st.header("Filtros")
        st.caption("Aguardando ficheiros de dados.")
    st.warning(f"⚠️ Não foram encontrados os ficheiros padrão no repositório para o período de {meses_pt[mes_selecionado]} {ano_corrente} ({sigla_mes}). Certifique-se de que os carregou para o GitHub ou use o painel de admin.")
else:
    tab_prod, tab_fact = st.tabs(["📊 Produtividade de Oficinas", "💶 Faturação e Orçamento Diário"])

    with tab_prod:
        xl = pd.ExcelFile(path_mecanicos)
        df_mecanicos = xl.parse("Mecânicos" if "Mecânicos" in xl.sheet_names else xl.sheet_names[0])
        df_hf = pd.read_excel(path_hf)
        
        df_ausencias = pd.DataFrame()
        if file_ausencias_up is not None:
            df_ausencias = pd.read_excel(file_ausencias_up)
        elif path_ausencias and isinstance(path_ausencias, str) and os.path.exists(path_ausencias):
            df_ausencias = pd.read_excel(path_ausencias)

        df_ausencias_grouped = pd.DataFrame(columns=["Técnico_clean", "Ausencias_calc"])
        if not df_ausencias.empty:
            col_desc_aus = df_ausencias.columns[1] if len(df_ausencias.columns) > 1 else None
            col_op_aus = next((c for c in ["Mecânico (Nome)", "Nome", "Técnico", "Tecnico", "Operário", "Operario"] if c in df_ausencias.columns), df_ausencias.columns[3] if len(df_ausencias.columns) > 3 else df_ausencias.select_dtypes(include=['object']).columns[0])
            col_val_aus = df_ausencias.select_dtypes(include=['number']).columns[-1]

            if col_desc_aus and col_val_aus:
                df_aus_filtrado = df_ausencias[df_ausencias[col_desc_aus].astype(str).str.strip().str.lower() != "presença"]
                df_aus_filtrado["Técnico_clean"] = df_aus_filtrado[col_op_aus].astype(str).str.strip().str.upper()
                df_ausencias_grouped = df_aus_filtrado.groupby("Técnico_clean", as_index=False)[col_val_aus].sum()
                df_ausencias_grouped.rename(columns={col_val_aus: "Ausencias_calc"}, inplace=True)

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
            if "ÉVORA" in nome_upper or "EVORA" in nome_upper:
                for o in oficinas_oficiais:
                    if "ÉVORA" in o.upper() or "EVORA" in o.upper(): return o
            if "SETÚBAL" in nome_upper or "SETUBAL" in nome_upper:
                for o in oficinas_oficiais:
                    if "SETÚBAL" in o.upper() or "SETUBAL" in o.upper(): return o
            if "SINES" in nome_upper:
                for o in oficinas_oficiais:
                    if "SINES" in o.upper(): return o
            return nome_str

        candidatos_op = ["Operário", "Operario", "Técnico", "Tecnico", "Nome", "Mecânico", "Mecanico"]
        col_operario_hf = next((c for c in candidatos_op if c in df_hf.columns), None)
        if not col_operario_hf:
            cols_obj = df_hf.select_dtypes(include=['object']).columns
            col_operario_hf = cols_obj[0] if len(cols_obj) > 0 else df_hf.columns[0]

        col_fact = next((c for c in ["Tempos facturados", "Tempos facturados ", "Horas facturadas", "Facturado", "Valor"] if c in df_hf.columns), df_hf.select_dtypes(include=['number']).columns[-1])
        col_oficina_hf = next((c for c in ["Oficina", "Escritório", "Escritorio", "Local"] if c in df_hf.columns), None)
        
        df_hf["Operario_clean"] = df_hf[col_operario_hf].astype(str).str.strip().str.upper()
        if col_oficina_hf:
            df_hf["Oficina_Normalizada"] = df_hf[col_oficina_hf].apply(normalizar_oficina)
        else:
            df_hf["Oficina_Normalizada"] = "Geral"

        df_hf_known = df_hf[df_hf["Operario_clean"].notna() & (df_hf["Operario_clean"] != "") & (df_hf["Operario_clean"] != "NAN")]
        df_hf_grouped = df_hf_known.groupby("Operario_clean", as_index=False)[col_fact].sum()
        df_hf_grouped.rename(columns={col_fact: "Horas facturadas", "Operario_clean": "Técnico_clean"}, inplace=True)

        df_merged = pd.merge(df_mecanicos, df_hf_grouped, on="Técnico_clean", how="left")
        if not df_ausencias_grouped.empty:
            df_merged = pd.merge(df_merged, df_ausencias_grouped, on="Técnico_clean", how="left")
            df_merged["Ausencias"] = df_merged["Ausencias_calc"].fillna(0.0)
        else:
            df_merged["Ausencias"] = 0.0

        df_merged["Horas facturadas"] = df_merged["Horas facturadas"].fillna(0.0)
        df_merged["Horas reais dia"] = df_merged["Horas reais dia"].fillna(7.5)
        df_merged["Horas disponiveis"] = (dias_decorridos * df_merged["Horas reais dia"]) - df_merged["Ausencias"]

        colunas_finais = [
            "Técnico", "
