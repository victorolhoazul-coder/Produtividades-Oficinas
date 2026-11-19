import pandas as pd
import streamlit as st
from datetime import datetime, date, timedelta
import os
import glob
import altair as alt
import unicodedata
import io
import csv

st.set_page_config(page_title="Relatório de Produtividade e Faturação - Oficinas", layout="wide")[cite: 14]

# ==========================================
# INICIALIZAÇÃO DO SESSION STATE
# ==========================================
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False[cite: 14]
if "utilizador" not in st.session_state:
    st.session_state.utilizador = ""[cite: 14]
if "rapv_permitido" not in st.session_state:
    st.session_state.rapv_permitido = ""[cite: 14]
if "acesso_total" not in st.session_state:
    st.session_state.acesso_total = False[cite: 14]
if "is_admin" not in st.session_state:
    st.session_state.is_admin = False[cite: 14]
if "agrupar_perfil" not in st.session_state:
    st.session_state.agrupar_perfil = False[cite: 14]
if "oficinas_agrupadas_perfil" not in st.session_state:
    st.session_state.oficinas_agrupadas_perfil = [][cite: 14]

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
""", unsafe_allow_html=True)[cite: 14]

col_titulo, col_logo = st.columns([4, 1])[cite: 14]
with col_titulo:
    st.title("Relatório de Oficinas - Grupo Motorpor")[cite: 14]
with col_logo:
    logo_path = "logo_motorpor.jpg"[cite: 14]
    if os.path.exists(logo_path):
        st.image(logo_path, width=150)[cite: 14]

meses_pt = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 
    5: "maio", 6: "junho", 7: "julho", 8: "agosto", 
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
}[cite: 14]

meses_sigla = {
    1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 
    5: "MAI", 6: "JUN", 7: "JUL", 8: "AGO", 
    9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"
}[cite: 14]

master_file = "Gestao_Oficinas_Master.xlsx"[cite: 14]
df_perfis = pd.DataFrame()[cite: 14]
df_agrup_master = pd.DataFrame()[cite: 14]
df_orc_master = pd.DataFrame()[cite: 14]

if os.path.exists(master_file):
    try:
        xl_master = pd.ExcelFile(master_file)[cite: 14]
        if "Perfis" in xl_master.sheet_names:
            df_perfis = xl_master.parse("Perfis")[cite: 14]
            df_perfis["Utilizador_clean"] = df_perfis["Utilizador"].astype(str).str.strip().str.lower()[cite: 14]
        if "Agrupamentos" in xl_master.sheet_names:
            df_agrup_master = xl_master.parse("Agrupamentos")[cite: 14]
        if "Orçamentos" in xl_master.sheet_names:
            df_orc_master = xl_master.parse("Orçamentos")[cite: 14]
    except Exception as e:
        st.error(f"Erro ao ler o ficheiro mestre: {e}")[cite: 14]

tem_permissao_upload = st.session_state.get("autenticado", False) and st.session_state.get("is_admin", False)[cite: 14]

st.sidebar.header("Navegação e Módulo")[cite: 14]
opcoes_modulos = [
    "📊 Produtividade de Oficinas", 
    "💶 Faturação e Orçamento Diário"
][cite: 14]

modulo_selecionado = st.sidebar.radio(
    "Escolha o Módulo:",
    options=opcoes_modulos
)[cite: 14]

st.sidebar.markdown("---")[cite: 14]
st.sidebar.header("Seleção de Período")[cite: 14]
ano_corrente = st.sidebar.selectbox("Ano:", options=[2025, 2026, 2027], index=1)[cite: 14]
mes_atual_num = datetime.now().month[cite: 14]

mes_selecionado = st.sidebar.selectbox(
    "Mês:", 
    options=list(meses_pt.keys()), 
    format_func=lambda x: meses_pt[x].capitalize(), 
    index=list(meses_pt.keys()).index(mes_atual_num) if (mes_atual_num in meses_pt and ano_corrente == 2026) else 10
)[cite: 14]

sigla_mes = meses_sigla.get(mes_selecionado, "NOV")[cite: 14]

cidade_feriado = st.sidebar.selectbox(
    "Local para Feriados Municipais:",
    options=["Nacional (Sem Municipal)", "Beja", "Évora", "Portimão", "Faro", "Setúbal", "Sines", "Portalegre"]
)[cite: 14]

if st.sidebar.button("Atualizar / Limpar Cache"):
    st.cache_data.clear()[cite: 14]
    st.rerun()[cite: 14]

ultima_atualizacao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")[cite: 14]
st.sidebar.caption(f"🔄 Última atualização: {ultima_atualizacao}")[cite: 14]

def obter_feriados_mes(ano, mes, cidade="Nacional (Sem Municipal)"):
    feriados = [
        date(ano, 1, 1), date(ano, 4, 25), date(ano, 5, 1),   
        date(ano, 6, 10), date(ano, 8, 15), date(ano, 10, 5),  
        date(ano, 12, 1), date(ano, 12, 8), date(ano, 12, 25), 
    ][cite: 14]
    if ano == 2026:
        feriados.extend([date(2026, 4, 3), date(2026, 6, 4)])[cite: 14]
        
    cidade_upper = str(cidade).upper()[cite: 14]
    if "BEJA" in cidade_upper: feriados.append(date(ano, 5, 14))[cite: 14]    
    elif "ÉVORA" in cidade_upper or "EVORA" in cidade_upper: feriados.append(date(ano, 6, 29))[cite: 14]    
    elif "PORTIMÃO" in cidade_upper or "PORTIMAO" in cidade_upper: feriados.append(date(ano, 12, 19))[cite: 14]   
    elif "FARO" in cidade_upper: feriados.append(date(ano, 9, 7))[cite: 14]     
    elif "SETÚBAL" in cidade_upper or "SETUBAL" in cidade_upper: feriados.append(date(ano, 9, 15))[cite: 14]    
    elif "SINES" in cidade_upper: feriados.append(date(ano, 11, 24))[cite: 14]
    elif "PORTALEGRE" in cidade_upper: feriados.append(date(ano, 5, 23))[cite: 14]
    return feriados[cite: 14]

def calcular_dias_uteis_ate_data(ano, mes, data_limite, cidade="Nacional (Sem Municipal)"):
    inicio = date(ano, mes, 1)[cite: 14]
    fim_mes = date(ano + 1, 1, 1) - timedelta(days=1) if mes == 12 else date(ano, mes + 1, 1) - timedelta(days=1)[cite: 14]
    limite = min(data_limite, fim_mes)[cite: 14]
    if limite < inicio: return 0[cite: 14]
    feriados = obter_feriados_mes(ano, mes, cidade)[cite: 14]
    dias_uteis = 0[cite: 14]
    atual = inicio[cite: 14]
    while atual <= limite:
        if atual.weekday() < 5 and atual not in feriados:
            dias_uteis += 1[cite: 14]
        atual += timedelta(days=1)[cite: 14]
    return max(1, dias_uteis)[cite: 14]

def calcular_dias_uteis_mes(ano, mes, cidade="Nacional (Sem Municipal)"):
    proximo_mes = date(ano + 1, 1, 1) if mes == 12 else date(ano, mes + 1, 1)[cite: 14]
    inicio = date(ano, mes, 1)[cite: 14]
    fim = proximo_mes - timedelta(days=1)[cite: 14]
    feriados = obter_feriados_mes(ano, mes, cidade)[cite: 14]
    dias_uteis = 0[cite: 14]
    atual = inicio[cite: 14]
    while atual <= fim:
        if atual.weekday() < 5 and atual not in feriados:
            dias_uteis += 1[cite: 14]
        atual += timedelta(days=1)[cite: 14]
    return dias_uteis[cite: 14]

dias_uteis_padrao_mes = calcular_dias_uteis_mes(ano_corrente, mes_selecionado, cidade_feriado)[cite: 14]
data_ontem = datetime.now().date() - timedelta(days=1)[cite: 14]
dias_decorridos_calc = calcular_dias_uteis_ate_data(ano_corrente, mes_selecionado, data_ontem, cidade_feriado)[cite: 14]
decorridos_iniciais = dias_uteis_padrao_mes if (mes_selecionado < mes_atual_num or ano_corrente < 2026) else min(dias_decorridos_calc, dias_uteis_padrao_mes)[cite: 14]

dias_uteis_padrao_mes = st.sidebar.slider(
    f"Total de Dias Úteis ({meses_pt[mes_selecionado].capitalize()}):", 
    min_value=15, max_value=25, value=dias_uteis_padrao_mes
)[cite: 14]

dias_decorridos = st.sidebar.number_input(
    "Dias úteis decorridos:", 
    min_value=1, max_value=dias_uteis_padrao_mes, 
    value=decorridos_iniciais if (mes_selecionado == mes_atual_num and ano_corrente == 2026) else (dias_uteis_padrao_mes if mes_selecionado <= mes_atual_num else 1)
)[cite: 14]

percentagem_progresso = (dias_decorridos / dias_uteis_padrao_mes) * 100 if dias_uteis_padrao_mes > 0 else 0[cite: 14]
progresso_decimal = min(max(dias_decorridos / dias_uteis_padrao_mes, 0.0), 1.0)[cite: 14]

st.sidebar.progress(
    progresso_decimal,
    text=f"Progresso: {dias_decorridos} de {dias_uteis_padrao_mes} dias ({percentagem_progresso:.1f}%)"
)[cite: 14]

st.sidebar.markdown("---")[cite: 14]
container_filtros = st.sidebar.container()[cite: 14]

if tem_permissao_upload:
    st.sidebar.markdown("---")[cite: 14]
    st.sidebar.header("Atualizar Ficheiros (Admin opcional)")[cite: 14]
    st.sidebar.file_uploader("Substituir 'tabela mecânicos'", type=["xlsx"], key="up_mec")[cite: 14]
    st.sidebar.file_uploader("Substituir 'Comparativo de Horas'", type=["xlsx"], key="up_hf")[cite: 14]
    st.sidebar.file_uploader("Substituir 'Listagem de Ausencias'", type=["xlsx"], key="up_aus")[cite: 14]
    st.sidebar.file_uploader("Mapa de Faturação 1 (CSV)", type=["csv"], key="up_f1")[cite: 14]
    st.sidebar.file_uploader("Mapa de Faturação 2 (CSV)", type=["csv"], key="up_f2")[cite: 14]

st.sidebar.markdown("---")[cite: 14]
container_login = st.sidebar.container()[cite: 14]
with container_login:
    st.header("Acesso / Login")[cite: 14]
    if not st.session_state.autenticado:
        if not df_perfis.empty:
            user_input = st.text_input("Utilizador:").strip().lower()[cite: 14]
            pass_input = st.text_input("Password:", type="password").strip()[cite: 14]

            if st.button("Entrar"):
                if user_input and pass_input:
                    match = df_perfis[df_perfis["Utilizador_clean"] == user_input][cite: 14]
                    if not match.empty:
                        st.session_state.autenticado = True[cite: 14]
                        st.session_state.utilizador = match.iloc[0]["Utilizador"][cite: 14]
                        rapv_perm = str(match.iloc[0]["RAPV Permitidos"]).strip()[cite: 14]
                        
                        st.session_state.rapv_permitido = rapv_perm[cite: 14]
                        st.session_state.is_admin = (rapv_perm.lower() == "administrador")[cite: 14]
                        st.session_state.acesso_total = (rapv_perm.lower() in ["administrador", "todas"])[cite: 14]
                        
                        agrupar_val = str(match.iloc[0].get("Agrupar", "não")).strip().lower()[cite: 14]
                        st.session_state.agrupar_perfil = (agrupar_val == "sim")[cite: 14]
                        
                        agrupamento_str = str(match.iloc[0].get("Agrupamento", "")).strip()[cite: 14]
                        if agrupamento_str and agrupamento_str.lower() != "nan" and agrupamento_str.lower() != "não":
                            st.session_state.oficinas_agrupadas_perfil = [o.strip().upper() for o in agrupamento_str.split(",") if o.strip()][cite: 14]
                        else:
                            st.session_state.oficinas_agrupadas_perfil = [][cite: 14]
                        st.rerun()[cite: 14]
                    else:
                        st.error("Utilizador ou password incorretos.")[cite: 14]
                else:
                    st.warning("Preencha o utilizador e a password.")[cite: 14]
        else:
            st.error("Ficheiro 'Gestao_Oficinas_Master.xlsx' sem aba de perfis.")[cite: 14]
    else:
        st.success(f"Sessão iniciada como: **{st.session_state.utilizador}**")[cite: 14]
        if st.button("Logout / Mudar Utilizador"):
            st.session_state.autenticado = False[cite: 14]
            st.session_state.utilizador = ""[cite: 14]
            st.session_state.rapv_permitido = ""[cite: 14]
            st.session_state.acesso_total = False[cite: 14]
            st.session_state.is_admin = False[cite: 14]
            st.session_state.agrupar_perfil = False[cite: 14]
            st.session_state.oficinas_agrupadas_perfil = [][cite: 14]
            st.rerun()[cite: 14]

file_mecanicos_up = st.session_state.get("up_mec")[cite: 14]
file_hf_up = st.session_state.get("up_hf")[cite: 14]
file_ausencias_up = st.session_state.get("up_aus")[cite: 14]
file_fact1_up = st.session_state.get("up_f1")[cite: 14]
file_fact2_up = st.session_state.get("up_f2")[cite: 14]

path_mecanicos = file_mecanicos_up if file_mecanicos_up is not None else ("tabela mecânicos.xlsx" if os.path.exists("tabela mecânicos.xlsx") else None)[cite: 14]

nome_arq_hf = f"Comparativo de Horas_{sigla_mes}_{ano_corrente}.xlsx"[cite: 14]
path_hf = file_hf_up if file_hf_up is not None else (nome_arq_hf if os.path.exists(nome_arq_hf) else None)[cite: 14]

nome_arq_aus = f"Listagem de AusenciasPresenças_{sigla_mes}_{ano_corrente}.xlsx"[cite: 14]
path_ausencias = file_ausencias_up if file_ausencias_up is not None else (nome_arq_aus if os.path.exists(nome_arq_aus) else None)[cite: 14]

def converter_valor_pt(val_str):
    if pd.isna(val_str): return 0.0[cite: 14]
    s = str(val_str).replace("€", "").strip()[cite: 14]
    if not s: return 0.0[cite: 14]
    s = s.replace("−", "-").replace("–", "-")[cite: 14]
    if "." in s and "," in s:
        if s.rfind('.') > s.rfind(','):
            s = s.replace(",", "")[cite: 14]
        else:
            s = s.replace(".", "").replace(",", ".")[cite: 14]
    elif "," in s:
        s = s.replace(",", ".")[cite: 14]
    try: return float(s)[cite: 14]
    except: return 0.0[cite: 14]

def converter_data_robusta(val_data, ano_alvo, mes_alvo):
    if pd.isna(val_data):
        return datetime(ano_alvo, mes_alvo, 1)[cite: 14]
    
    s = str(val_data).strip()[cite: 14]
    if not s or s.lower() == "nan":
        return datetime(ano_alvo, mes_alvo, 1)[cite: 14]

    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%y", "%d-%m-%y"):
        try:
            dt = datetime.strptime(s[:10], fmt)[cite: 14]
            if dt.year < 2000:
                dt = dt.replace(year=dt.year + 2000)[cite: 14]
            return dt[cite: 14]
        except:
            continue[cite: 14]
            
    dt_pd = pd.to_datetime(s, dayfirst=True, errors="coerce")[cite: 14]
    if pd.notna(dt_pd):
        return dt_pd.to_pydatetime()[cite: 14]
        
    dt_pd2 = pd.to_datetime(s, dayfirst=False, errors="coerce")[cite: 14]
    if pd.notna(dt_pd2):
        return dt_pd2.to_pydatetime()[cite: 14]
        
    return datetime(ano_alvo, mes_alvo, 1)[cite: 14]

def limpar_str(s):
    if pd.isna(s): return ""[cite: 14]
    s_nfkd = unicodedata.normalize('NFKD', str(s))[cite: 14]
    return "".join([c for c in s_nfkd if not unicodedata.combining(c)]).upper().strip().rstrip(".")[cite: 14]

def limpar_prefixo_oficina(nome):
    limpo = str(nome).strip()[cite: 14]
    if limpo.upper().startswith("OF. "): limpo = limpo[4:].strip()[cite: 14]
    elif limpo.upper().startswith("OF "): limpo = limpo[3:].strip()[cite: 14]
    return limpo.rstrip(".")[cite: 14]

def oficina_permitida_check(oficina_nome):
    if st.session_state.acesso_total: return True[cite: 14]
    oficina_up = limpar_str(oficina_nome)[cite: 14]
    if oficina_up == "SEM PRODUTIVO": return True[cite: 14]
    
    for permitida in st.session_state.oficinas_agrupadas_perfil:
        p_clean = limpar_str(permitida)[cite: 14]
        p_clean_sem_of = limpar_prefixo_oficina(p_clean)[cite: 14]
        oficina_sem_of = limpar_prefixo_oficina(oficina_up)[cite: 14]
        if (p_clean == oficina_up or 
            p_clean_sem_of == oficina_sem_of or 
            p_clean in oficina_up or 
            oficina_up in p_clean or
            p_clean_sem_of in oficina_sem_of or
            oficina_sem_of in p_clean_sem_of):
            return True[cite: 14]
    return False[cite: 14]

def carregar_dados_csv_faturacao():
    df_fact_total = pd.DataFrame()[cite: 14]
    csv_fontes = [file_fact1_up, file_fact2_up][cite: 14]
    
    if all(f is None for f in csv_fontes):
        encontrados = [][cite: 14]
        dir_base = os.path.dirname(os.path.abspath(__file__)) if "__file__" in locals() else "."[cite: 14]
        
        for d in [dir_base, "."]:
            for ext in ["*.csv", "*.CSV"]:
                encontrados.extend(glob.glob(os.path.join(d, ext)))[cite: 14]
        
        fich_exatos = [
            "MOTOREX_MAPA_FACTURAÇÃO__2026_11.csv",
            "MULTIAUTO_MAPA_FACTURAÇÃO__2026_11.csv"
        ][cite: 14]
        for fe in fich_exatos:
            caminho_fe = os.path.join(dir_base, fe)[cite: 14]
            if os.path.exists(caminho_fe):
                encontrados.append(caminho_fe)[cite: 14]
            elif os.path.exists(fe):
                encontrados.append(fe)[cite: 14]

        encontrados = list(set(encontrados))[cite: 14]
        st.sidebar.info(f"📁 CSVs detetados: {encontrados}")[cite: 14]
        
        for arq in encontrados:
            csv_fontes.append(arq)[cite: 14]

    for file_f in csv_fontes:
        if file_f is not None:
            try:
                if hasattr(file_f, "getvalue"):
                    conteudo = file_f.getvalue().decode("iso-8859-1", errors="ignore")[cite: 14]
                elif hasattr(file_f, "read"):
                    file_f.seek(0)[cite: 14]
                    res = file_f.read()[cite: 14]
                    conteudo = res.decode("iso-8859-1", errors="ignore") if isinstance(res, bytes) else res[cite: 14]
                else:
                    with open(file_f, "r", encoding="iso-8859-1", errors="ignore") as f:
                        conteudo = f.read()[cite: 14]
                
                linhas_validas = [][cite: 14]
                for sep in [";", "\t", ","]:
                    f_io = io.StringIO(conteudo)[cite: 14]
                    reader = csv.reader(f_io, delimiter=sep)[cite: 14]
                    temp_linhas = [row for row in reader if len(row) > 2][cite: 14]
                    if len(temp_linhas) > 5:
                        linhas_validas = temp_linhas[cite: 14]
                        break[cite: 14]
                
                if not linhas_validas:
                    f_io = io.StringIO(conteudo)[cite: 14]
                    reader = csv.reader(f_io, delimiter=";")[cite: 14]
                    linhas_validas = [row for row in reader if len(row) > 0][cite: 14]

                if linhas_validas:
                    df_temp = pd.DataFrame(linhas_validas)[cite: 14]
                    df_temp["Origem_CSV"] = getattr(file_f, "name", str(file_f))[cite: 14]
                    df_fact_total = pd.concat([df_fact_total, df_temp], ignore_index=True)[cite: 14]
            except Exception as e:
                st.error(f"Erro ao ler CSV: {e}")[cite: 14]
    return df_fact_total[cite: 14]

# ==========================================
# CORPO PRINCIPAL
# ==========================================
if not st.session_state.autenticado:
    st.warning("🔒 Por favor, faça login na barra lateral para aceder aos relatórios.")[cite: 14]
elif not path_mecanicos or not path_hf:
    with container_filtros:
        st.header("Filtros")[cite: 14]
        st.caption("Aguardando ficheiros de dados.")[cite: 14]
    st.warning(f"⚠️ Não foram encontrados os ficheiros padrão no disco para o período de {meses_pt[mes_selecionado].capitalize()} {ano_corrente} ({sigla_mes}). Certifique-se de que os ficheiros estão na pasta do projeto.")[cite: 14]
else:
    if modulo_selecionado == "📊 Produtividade de Oficinas":
        xl = pd.ExcelFile(path_mecanicos)[cite: 14]
        df_mecanicos = xl.parse("Mecânicos" if "Mecânicos" in xl.sheet_names else xl.sheet_names[0])[cite: 14]
        df_hf = pd.read_excel(path_hf)[cite: 14]
        
        df_ausencias = pd.DataFrame()[cite: 14]
        if path_ausencias is not None:
            try:
                df_ausencias = pd.read_excel(path_ausencias)[cite: 14]
            except:
                pass[cite: 14]

        df_ausencias_grouped = pd.DataFrame(columns=["Técnico_clean", "Ausencias_calc"])[cite: 14]
        if not df_ausencias.empty:
            col_data_aus = next((c for c in ["Data", "Data (Data)", "Dia"] if c in df_ausencias.columns), df_ausencias.columns[4] if len(df_ausencias.columns) > 4 else None)[cite: 14]
            col_desc_aus = df_ausencias.columns[1] if len(df_ausencias.columns) > 1 else None[cite: 14]
            col_op_aus = next((c for c in ["Mecânico (Nome)", "Nome", "Técnico", "Tecnico", "Operário", "Operario"] if c in df_ausencias.columns), df_ausencias.columns[3] if len(df_ausencias.columns) > 3 else df_ausencias.select_dtypes(include=['object']).columns[0])[cite: 14]
            col_val_aus = df_ausencias.select_dtypes(include=['number']).columns[-1][cite: 14]

            if col_desc_aus and col_val_aus:
                df_aus_filtrado = df_ausencias.copy()[cite: 14]
                if col_data_aus:
                    df_aus_filtrado["Data_Parsed"] = pd.to_datetime(df_aus_filtrado[col_data_aus], dayfirst=True, errors="coerce")[cite: 14]
                    if ano_corrente == datetime.now().year and mes_selecionado == datetime.now().month:
                        df_aus_filtrado = df_aus_filtrado[df_aus_filtrado["Data_Parsed"].dt.date <= data_ontem][cite: 14]

                df_aus_filtrado = df_aus_filtrado[df_aus_filtrado[col_desc_aus].astype(str).str.strip().str.lower() != "presença"][cite: 14]
                df_aus_filtrado["Técnico_clean"] = df_aus_filtrado[col_op_aus].astype(str).str.strip().str.upper()[cite: 14]
                df_ausencias_grouped = df_aus_filtrado.groupby("Técnico_clean", as_index=False)[col_val_aus].sum()[cite: 14]
                df_ausencias_grouped.rename(columns={col_val_aus: "Ausencias_calc"}, inplace=True)[cite: 14]

        df_mecanicos.rename(columns={"Técnicos": "Técnico", "Nome": "Técnico"}, inplace=True)[cite: 14]
        df_mecanicos["Técnico_clean"] = df_mecanicos["Técnico"].astype(str).str.strip().str.upper()[cite: 14]
        df_mecanicos["Oficina"] = df_mecanicos["Oficina"].astype(str).str.strip()[cite: 14]
        df_mecanicos["Local"] = df_mecanicos["Local"].astype(str).str.strip()[cite: 14]

        mapa_prod_master = {}[cite: 14]
        if not df_agrup_master.empty:
            for _, r in df_agrup_master.iterrows():
                ofic_orig = str(r.iloc[0]).strip()[cite: 14]
                agrup_prod = str(r.iloc[5]).strip() if len(r) > 5 and pd.notna(r.iloc[5]) else ofic_orig[cite: 14]
                if agrup_prod and agrup_prod.lower() != "nan":
                    mapa_prod_master[limpar_str(ofic_orig)] = agrup_prod[cite: 14]

        def aplicar_agrupamento_prod(nome_ofic):
            u = limpar_str(nome_ofic)[cite: 14]
            return mapa_prod_master.get(u, nome_ofic)[cite: 14]

        df_mecanicos["Oficina_Original"] = df_mecanicos["Oficina"][cite: 14]
        df_mecanicos["Oficina"] = df_mecanicos["Oficina"].apply(aplicar_agrupamento_prod)[cite: 14]

        oficinas_oficiais = df_mecanicos["Oficina"].unique().tolist()[cite: 14]
        dict_oficina_local = dict(zip(df_mecanicos["Oficina"].str.upper(), df_mecanicos["Local"]))[cite: 14]

        def normalizar_oficina_e_local(nome_oficina):
            if pd.isna(nome_oficina): return "Geral", "Sem produtivo"[cite: 14]
            nome_str = str(nome_oficina).strip()[cite: 14]
            agrupado = aplicar_agrupamento_prod(nome_str)[cite: 14]
            for ofic_oficial in oficinas_oficiais:
                if ofic_oficial.upper() == agrupado.upper():
                    return ofic_oficial, dict_oficina_local.get(ofic_oficial.upper(), "Geral")[cite: 14]
            return agrupado, "Sem produtivo"[cite: 14]

        candidatos_op = ["Operário", "Operario", "Técnico", "Tecnico", "Nome", "Mecânico", "Mecanico"][cite: 14]
        col_operario_hf = next((c for c in candidatos_op if c in df_hf.columns), df_hf.select_dtypes(include=['object']).columns[0])[cite: 14]
        col_fact = next((c for c in ["Tempos facturados", "Tempos facturados ", "Horas facturadas", "Facturado", "Valor"] if c in df_hf.columns), df_hf.select_dtypes(include=['number']).columns[-1])[cite: 14]
        col_oficina_hf = next((c for c in ["Oficina", "Escritório", "Escritorio", "Local"] if c in df_hf.columns), None)[cite: 14]
        
        df_hf["Operario_clean"] = df_hf[col_operario_hf].astype(str).str.strip().str.upper()[cite: 14]
        
        if col_oficina_hf:
            res_norm = df_hf[col_oficina_hf].apply(normalizar_oficina_e_local)[cite: 14]
            df_hf["Oficina_Normalizada"] = [r[0] for r in res_norm][cite: 14]
            df_hf["Local_Normalizado"] = [r[1] for r in res_norm][cite: 14]
        else:
            df_hf["Oficina_Normalizada"] = "Geral"[cite: 14]
            df_hf["Local_Normalizado"] = "Sem produtivo"[cite: 14]

        df_hf_known = df_hf[df_hf["Operario_clean"].notna() & (df_hf["Operario_clean"] != "") & (df_hf["Operario_clean"] != "NAN")].copy()[cite: 14]
        df_hf_sem_op = df_hf[df_hf["Operario_clean"].isna() | (df_hf["Operario_clean"] == "") | (df_hf["Operario_clean"] == "NAN")].copy()[cite: 14]

        df_hf_grouped = df_hf_known.groupby("Operario_clean", as_index=False)[col_fact].sum()[cite: 14]
        df_hf_grouped.rename(columns={col_fact: "Horas facturadas", "Operario_clean": "Técnico_clean"}, inplace=True)[cite: 14]

        df_merged = pd.merge(df_mecanicos, df_hf_grouped, on="Técnico_clean", how="left")[cite: 14]
        if not df_ausencias_grouped.empty:
            df_merged = pd.merge(df_merged, df_ausencias_grouped, on="Técnico_clean", how="left")[cite: 14]
            df_merged["Ausencias"] = df_merged["Ausencias_calc"].fillna(0.0)[cite: 14]
        else:
            df_merged["Ausencias"] = 0.0[cite: 14]

        df_merged["Horas facturadas"] = df_merged["Horas facturadas"].fillna(0.0)[cite: 14]
        df_merged["Horas potencial dia"] = dias_decorridos * df_merged["Horas reais dia"][cite: 14]
        df_merged["Horas disponiveis"] = df_merged["Horas potencial dia"] - df_merged["Ausencias"][cite: 14]

        colunas_finais = ["Técnico", "Local", "Oficina", "Função", "Horas potencial dia", "Horas reais dia", "Ausencias", "Horas disponiveis", "Horas facturadas", "Produtividade %"][cite: 14]
        
        df_merged["Produtividade %"] = 0.0[cite: 14]
        mask_disp = df_merged["Horas disponiveis"] > 0[cite: 14]
        df_merged.loc[mask_disp, "Produtividade %"] = (df_merged.loc[mask_disp, "Horas facturadas"] / df_merged.loc[mask_disp, "Horas disponiveis"]) * 100[cite: 14]

        df_exibicao_tecnicos = df_merged[[c for c in colunas_finais if c in df_merged.columns]].copy()[cite: 14]

        linhas_sem_produtivo = [][cite: 14]
        if not df_hf_sem_op.empty:
            df_sem_op_grouped = df_hf_sem_op.groupby(["Local_Normalizado", "Oficina_Normalizada"], as_index=False)[col_fact].sum()[cite: 14]
            for _, row in df_sem_op_grouped.iterrows():
                if row[col_fact] > 0:
                    linhas_sem_produtivo.append({
                        "Técnico": "Sem produtivo associado", "Local": row["Local_Normalizado"], "Oficina": row["Oficina_Normalizada"],
                        "Função": "Sem Produtivo", "Horas potencial dia": 0.0, "Horas reais dia": 0.0, "Ausencias": 0.0,
                        "Horas disponiveis": 0.0, "Horas facturadas": row[col_fact], "Produtividade %": 0.0
                    })[cite: 14]

        df_sem_prod_final = pd.DataFrame(linhas_sem_produtivo)[cite: 14]
        df_completo = pd.concat([df_exibicao_tecnicos, df_sem_prod_final], ignore_index=True) if not df_sem_prod_final.empty else df_exibicao_tecnicos[cite: 14]

        def checar_perm_prod(oficina_nome):
            if st.session_state.acesso_total: return True[cite: 14]
            if oficina_permitida_check(oficina_nome): return True[cite: 14]
            
            oficina_up = limpar_str(oficina_nome)[cite: 14]
            user_rapv = limpar_str(st.session_state.rapv_permitido)[cite: 14]
            if not df_agrup_master.empty:
                for _, r in df_agrup_master.iterrows():
                    agrup_p = limpar_str(r.iloc[5] if len(r) > 5 and pd.notna(r.iloc[5]) else r.iloc[0])[cite: 14]
                    rapv_m = limpar_str(r.iloc[4] if len(r) > 4 and pd.notna(r.iloc[4]) else "")[cite: 14]
                    if rapv_m == user_rapv and (agrup_p == oficina_up or oficina_up in agrup_p or agrup_p in oficina_up or limpar_prefixo_oficina(agrup_p) in limpar_prefixo_oficina(oficina_up)):
                        return True[cite: 14]
            return False[cite: 14]

        df_completo = df_completo[df_completo["Oficina"].apply(checar_perm_prod)][cite: 14]

        with container_filtros:
            st.header("Filtros")[cite: 14]
            locais_unicos = sorted(df_completo["Local"].dropna().astype(str).unique().tolist())[cite: 14]
            locais_disponiveis = (["Todas"] + locais_unicos) if len(locais_unicos) > 1 else locais_unicos[cite: 14]
            sel_local = st.selectbox("Selecionar Cidade (Local)", locais_disponiveis if locais_disponiveis else ["Sem acesso"])[cite: 14]

            df_filtrado_local = df_completo[df_completo["Local"] == sel_local] if sel_local != "Todas" else df_completo[cite: 14]

            oficinas_unicas = sorted(df_filtrado_local["Oficina"].dropna().astype(str).unique().tolist())[cite: 14]
            oficinas_disponiveis = (["Todas"] + oficinas_unicas) if len(oficinas_unicas) > 1 else oficinas_unicas[cite: 14]
            sel_oficina = st.selectbox("Selecionar Oficina", oficinas_disponiveis if oficinas_disponiveis else ["Sem acesso"])[cite: 14]

        df_final_filtro = df_filtrado_local[df_filtrado_local["Oficina"] == sel_oficina] if sel_oficina != "Todas" else df_filtrado_local[cite: 14]

        tot_h_potencial = df_final_filtro["Horas potencial dia"].sum()[cite: 14]
        tot_h_reais = df_final_filtro["Horas reais dia"].sum()[cite: 14]
        tot_ausencias = df_final_filtro["Ausencias"].sum()[cite: 14]
        tot_h_disp = df_final_filtro["Horas disponiveis"].sum()[cite: 14]
        tot_h_fact = df_final_filtro["Horas facturadas"].sum()[cite: 14]
        tot_prod = (tot_h_fact / tot_h_disp * 100) if tot_h_disp > 0 else 0.0[cite: 14]

        linha_total = {
            "Técnico": "TOTAL", "Local": "-", "Oficina": "-", "Função": "-",
            "Horas potencial dia": tot_h_potencial, "Horas reais dia": tot_h_reais, "Ausencias": tot_ausencias,
            "Horas disponiveis": tot_h_disp, "Horas facturadas": tot_h_fact, "Produtividade %": tot_prod
        }[cite: 14]

        df_exibicao_final = pd.concat([df_final_filtro, pd.DataFrame([linha_total])], ignore_index=True)[cite: 14]

        def destacar_total(s):
            is_total = s.index.isin(df_exibicao_final[df_exibicao_final["Técnico"] == "TOTAL"].index)[cite: 14]
            return ['font-weight: bold; background-color: #f0f2f6;' if v else '' for v in is_total][cite: 14]

        st.subheader(f"Mapa de Produtividade - {meses_pt[mes_selecionado].capitalize()} {ano_corrente}")[cite: 14]
        st.dataframe(
            df_exibicao_final.style.apply(destacar_total, axis=1).format({
                "Horas potencial dia": "{:.1f}", "Horas reais dia": "{:.1f}", "Ausencias": "{:.1f}",
                "Horas disponiveis": "{:.1f}", "Horas facturadas": "{:.2f}", "Produtividade %": "{:.1f}%"
            }, na_rep="-"),
            use_container_width=True,
            hide_index=True,
            height=800
        )[cite: 14]

    elif modulo_selecionado == "💶 Faturação e Orçamento Diário":
        st.subheader(f"Faturação e Orçamento Diário - {meses_pt[mes_selecionado].capitalize()} {ano_corrente}")[cite: 14]
        
        df_orc = df_orc_master.copy()[cite: 14]
        df_fact_total = carregar_dados_csv_faturacao()[cite: 14]
                
        if df_fact_total.empty:
            st.warning("⚠️ Não foram encontrados ficheiros CSV de faturação no disco.")[cite: 14]
        else:
            num_cols = df_fact_total.shape[1][cite: 14]
            idx_ofic = 0[cite: 14]
            idx_val = 24 if num_cols > 24 else num_cols - 2[cite: 14]
            idx_data = 28 if num_cols > 28 else num_cols - 1[cite: 14]

            df_fact_total["Oficina_CSV"] = df_fact_total.iloc[:, idx_ofic].astype(str).str.strip()[cite: 14]
            invalidos_csv = ["NAN", "", "TOTAL", "MAPA DE FACTURAÇÃO", "OFICINA", "OFICINAS", "NOME", "SECCAO"][cite: 14]
            df_fact_total = df_fact_total[~df_fact_total["Oficina_CSV"].str.upper().isin(invalidos_csv)][cite: 14]
            df_fact_total = df_fact_total[~df_fact_total["Oficina_CSV"].str.startswith("*")][cite: 14]

            df_fact_total["Data_Parsed"] = df_fact_total.iloc[:, idx_data].apply(lambda x: converter_data_robusta(x, ano_corrente, mes_selecionado))[cite: 14]
            df_fact_total["Valor_Num"] = df_fact_total.iloc[:, idx_val].apply(converter_valor_pt)[cite: 14]

            mapa_fact_master = {}[cite: 14]
            mapa_rapv_master = {}[cite: 14]
            lista_suboficinas_por_grupo_fact = {}[cite: 14]

            if not df_agrup_master.empty:
                for _, r in df_agrup_master.iterrows():
                    ofic_orig = str(r.iloc[0]).strip()[cite: 14]
                    agrup_fact = str(r.iloc[6]).strip() if len(r) > 6 and pd.notna(r.iloc[6]) else ofic_orig[cite: 14]
                    rapv_val = str(r.iloc[4]).strip() if len(r) > 4 and pd.notna(r.iloc[4]) else "Geral"[cite: 14]
                    
                    if not agrup_fact or agrup_fact.lower() == "nan":
                        agrup_fact = ofic_orig[cite: 14]

                    ofic_clean = limpar_str(ofic_orig)[cite: 14]
                    mapa_fact_master[ofic_clean] = agrup_fact[cite: 14]
                    mapa_rapv_master[agrup_fact] = rapv_val[cite: 14]

                    if agrup_fact not in lista_suboficinas_por_grupo_fact:
                        lista_suboficinas_por_grupo_fact[agrup_fact] = set()[cite: 14]
                    lista_suboficinas_por_grupo_fact[agrup_fact].add(ofic_clean)[cite: 14]

            def obter_grupo_fact(nome_csv):
                u = limpar_str(nome_csv)[cite: 14]
                if u in mapa_fact_master:
                    return mapa_fact_master[u][cite: 14]
                
                is_colisao = "COLISÃO" in u or "COLISAO" in u[cite: 14]
                melhor_match = None[cite: 14]
                max_tokens = 0[cite: 14]
                
                for chave_master, agrup in mapa_fact_master.items():
                    agrup_is_colisao = "COLISÃO" in agrup.upper() or "COLISAO" in agrup.upper()[cite: 14]
                    if is_colisao != agrup_is_colisao:
                        continue[cite: 14]
                        
                    tokens_master = [t for t in chave_master.replace("OF.", "").replace("COLISÃO", "").replace("COLISAO", "").split() if t not in ["DE", "DA", "E"]][cite: 14]
                    if tokens_master and all(t in u for t in tokens_master):
                        if len(tokens_master) > max_tokens:
                            max_tokens = len(tokens_master)[cite: 14]
                            melhor_match = agrup[cite: 14]
                            
                if melhor_match:
                    return melhor_match[cite: 14]
                    
                for agrup in mapa_fact_master.values():
                    if limpar_str(agrup) in u or u in limpar_str(agrup):
                        return agrup[cite: 14]
                        
                return nome_csv[cite: 14]

            df_fact_total["Oficina_Agrupada"] = df_fact_total["Oficina_CSV"].apply(obter_grupo_fact)[cite: 14]
            df_fact_total["RAPV"] = df_fact_total["Oficina_Agrupada"].map(lambda g: mapa_rapv_master.get(g, "Geral"))[cite: 14]

            oficinas_disponiveis_fat = sorted(list(set(df_fact_total["Oficina_Agrupada"].dropna().astype(str).unique())))[cite: 14]

            if st.session_state.acesso_total:
                oficinas_autorizadas_lista = oficinas_disponiveis_fat[cite: 14]
            else:
                user_rapv = str(st.session_state.rapv_permitido).strip().upper()[cite: 14]
                oficinas_autorizadas_lista = [
                    o for o in oficinas_disponiveis_fat 
                    if mapa_rapv_master.get(o, "").strip().upper() == user_rapv or oficina_permitida_check(o)
                ][cite: 14]
            if not oficinas_autorizadas_lista: oficinas_autorizadas_lista = ["Sem acesso"][cite: 14]

            col_sel_rapv, col_sel_ofic = st.columns(2)[cite: 14]
            
            if st.session_state.acesso_total:
                rapvs_autorizados = sorted(list(set(mapa_rapv_master.values())))[cite: 14]
            else:
                rapv_usuario_atual = str(st.session_state.rapv_permitido).strip()[cite: 14]
                rapvs_autorizados = [rapv_usuario_atual] if rapv_usuario_atual else ["Geral"][cite: 14]

            with col_sel_rapv:
                rapv_selecionado = st.selectbox("Selecionar RAPV:", options=rapvs_autorizados)[cite: 14]

            oficinas_filtradas_rapv = [
                o for o in oficinas_autorizadas_lista 
                if mapa_rapv_master.get(o, "").strip().upper() == rapv_selecionado.strip().upper()
            ][cite: 14]
            if not oficinas_filtradas_rapv: oficinas_filtradas_rapv = [oficinas_autorizadas_lista[0]][cite: 14]

            with col_sel_ofic:
                oficina_selecionada_fat = st.selectbox("Selecionar Oficina (Faturação):", options=oficinas_filtradas_rapv)[cite: 14]
                
            sel_clean = limpar_str(oficina_selecionada_fat)[cite: 14]
            suboficinas_grupo = lista_suboficinas_por_grupo_fact.get(oficina_selecionada_fat, {sel_clean})[cite: 14]
            
            df_oficina_fat = df_fact_total[
                (df_fact_total["Oficina_Agrupada"].apply(limpar_str) == sel_clean) |
                (df_fact_total["Oficina_CSV"].apply(limpar_str).isin(suboficinas_grupo))
            ][cite: 14]
            
            df_oficina_fat = df_oficina_fat[(df_oficina_fat["Data_Parsed"].dt.month == mes_selecionado) & (df_oficina_fat["Data_Parsed"].dt.year == ano_corrente)][cite: 14]
            
            df_diario = df_oficina_fat.groupby(df_oficina_fat["Data_Parsed"].dt.date, as_index=False)["Valor_Num"].sum()[cite: 14]
            df_diario.rename(columns={"Valor_Num": "Facturacao_Diaria", "Data_Parsed": "Data"}, inplace=True)[cite: 14]
            df_diario = df_diario.dropna(subset=["Data"]).sort_values("Data")[cite: 14]
            
            inicio_mes = date(ano_corrente, mes_selecionado, 1)[cite: 14]
            fim_mes = date(ano_corrente + 1, 1, 1) - timedelta(days=1) if mes_selecionado == 12 else date(ano_corrente, mes_selecionado + 1, 1) - timedelta(days=1)[cite: 14]
            feriados_mes = obter_feriados_mes(ano_corrente, mes_selecionado, cidade_feriado)[cite: 14]
            
            dias_uteis_lista = [][cite: 14]
            atual = inicio_mes[cite: 14]
            while atual <= fim_mes:
                if atual.weekday() < 5 and atual not in feriados_mes: dias_uteis_lista.append(atual)[cite: 14]
                atual += timedelta(days=1)[cite: 14]
                
            df_completo_dias = pd.DataFrame({"Data": dias_uteis_lista})[cite: 14]
            df_completo_dias = pd.merge(df_completo_dias, df_diario, on="Data", how="left").fillna(0.0)[cite: 14]
            df_completo_dias["Data_Str"] = pd.to_datetime(df_completo_dias["Data"]).dt.strftime("%d/%m/%Y")[cite: 14]
            
            hoje_atual = datetime.now().date()[cite: 14]
            data_limite_grafico = hoje_atual if (ano_corrente == hoje_atual.year and mes_selecionado == hoje_atual.month) else fim_mes[cite: 14]

            df_completo_dias["Facturacao_Acumulada"] = df_completo_dias["Facturacao_Diaria"].cumsum()[cite: 14]
            df_completo_dias["Facturacao_Acumulada_Graf"] = df_completo_dias.apply(lambda r: r["Facturacao_Acumulada"] if r["Data"] <= data_limite_grafico else None, axis=1)[cite: 14]
            df_completo_dias["Facturacao_Diaria_Graf"] = df_completo_dias.apply(lambda r: r["Facturacao_Diaria"] if r["Data"] <= data_limite_grafico else 0.0, axis=1)[cite: 14]

            orcamento_mensal = 0.0[cite: 14]
            if not df_orc.empty:
                col_ofic = df_orc.columns[0][cite: 14]
                col_mes = df_orc.columns[3][cite: 14]
                col_val = df_orc.columns[4][cite: 14]
                nome_mes_alvo = limpar_str(meses_pt[mes_selecionado])[cite: 14]
                
                for _, row in df_orc.iterrows():
                    ofic_raw = limpar_prefixo_oficina(str(row[col_ofic]))[cite: 14]
                    ofic_row = limpar_str(ofic_raw)[cite: 14]
                    if not ofic_row or any(exc in ofic_row for exc in ["TOTAL", "GRUPO", "MAPA", "ACUMULADO"]): continue[cite: 14]
                    
                    val_mes_cell = row[col_mes][cite: 14]
                    mes_encontrado = False[cite: 14]
                    if isinstance(val_mes_cell, (pd.Timestamp, datetime, date)):
                        if val_mes_cell.month == mes_selecionado: mes_encontrado = True[cite: 14]
                    else:
                        mes_row_str = limpar_str(val_mes_cell)[cite: 14]
                        if nome_mes_alvo in mes_row_str or mes_row_str in nome_mes_alvo: mes_encontrado = True[cite: 14]
                        else:
                            try:
                                if int(val_mes_cell) == mes_selecionado: mes_encontrado = True[cite: 14]
                            except: pass[cite: 14]
                            
                    if mes_encontrado:
                        if sel_clean == ofic_row or sel_clean in ofic_row or ofic_row in sel_clean:
                            orcamento_mensal += converter_valor_pt(row[col_val])[cite: 14]
                                
            orcamento_diario = orcamento_mensal / dias_uteis_padrao_mes if dias_uteis_padrao_mes > 0 else 0.0[cite: 14]
            
            orcamentos_ac = [][cite: 14]
            acum = 0.0[cite: 14]
            for d in df_completo_dias["Data"]:
                acum += orcamento_diario[cite: 14]
                orcamentos_ac.append(acum)[cite: 14]
            df_completo_dias["Orcamento_Acumulado"] = orcamentos_ac[cite: 14]
            
            val_fat_reais = df_diario[df_diario["Data"] <= data_limite_grafico]["Facturacao_Diaria"][cite: 14]
            facturacao_a_data = val_fat_reais.sum() if not val_fat_reais.empty else 0.0[cite: 14]
            
            dias_uteis_decorridos_reais = len(df_completo_dias[df_completo_dias["Data"] <= data_limite_grafico])[cite: 14]
            orcamento_a_data = dias_uteis_decorridos_reais * orcamento_diario[cite: 14]
            desvio_orcamento = facturacao_a_data - orcamento_a_data[cite: 14]
            
            col_graf, col_kpi = st.columns([3, 1])[cite: 14]
            
            with col_graf:
                base = alt.Chart(df_completo_dias).encode(
                    x=alt.X('Data_Str:N', title='Data', sort=list(df_completo_dias["Data_Str"]), axis=alt.Axis(labelAngle=-45))
                )[cite: 14]
                barras = base.mark_bar(color='#115e59').encode(
                    y=alt.Y('Facturacao_Acumulada_Graf:Q', title='Valor (€)'),
                    tooltip=[
                        alt.Tooltip('Data_Str:N', title='Data'),
                        alt.Tooltip('Facturacao_Diaria_Graf:Q', title='Faturação Diária', format=',.2f'),
                        alt.Tooltip('Facturacao_Acumulada_Graf:Q', title='Faturação Acumulada', format=',.2f')
                    ]
                )[cite: 14]
                linha_orc = base.mark_line(color='#f97316', strokeWidth=3).encode(
                    y=alt.Y('Orcamento_Acumulado:Q', title='Orçamento Acumulado'),
                    tooltip=[
                        alt.Tooltip('Data_Str:N', title='Data'),
                        alt.Tooltip('Orcamento_Acumulado:Q', title='Orçamento Acumulado', format=',.2f')
                    ]
                )[cite: 14]
                chart = (barras + linha_orc).properties(height=420).interactive()[cite: 14]
                st.altair_chart(chart, use_container_width=True)[cite: 14]
                
            with col_kpi:
                orc_mensal_txt = f"{orcamento_mensal:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")[cite: 14]
                orc_diario_txt = f"{orcamento_diario:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")[cite: 14]
                fat_data_txt = f"{facturacao_a_data:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")[cite: 14]
                desvio_txt = f"{desvio_orcamento:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")[cite: 14]
                cor_desvio = "green" if desvio_orcamento >= 0 else "red"[cite: 14]

                st.metric(label="ORÇAMENTO MENSAL", value=orc_mensal_txt)[cite: 14]
                st.metric(label="ORÇAMENTO DIÁRIO", value=orc_diario_txt)[cite: 14]
                st.metric(label="FATURAÇÃO À DATA", value=fat_data_txt)[cite: 14]
                
                st.markdown(f"""
                    <div style="font-size: 14px; font-weight: 600; color: rgb(49, 51, 63); margin-bottom: 0px;">DESVIO ORÇAMENTO</div>
                    <div style="font-size: 24px; font-weight: 700; color: {cor_desvio}; margin-top: 0px; margin-bottom: 0px;">{desvio_txt}</div>
                """, unsafe_allow_html=True)[cite: 14]
                st.metric(label="", value="", delta=desvio_txt, label_visibility="collapsed")[cite: 14]