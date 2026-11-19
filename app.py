import pandas as pd
import streamlit as st
from datetime import datetime, date, timedelta
import os
import glob
import altair as alt
import unicodedata
import io
import csv

st.set_page_config(page_title="Relatório de Produtividade e Faturação - Oficinas", layout="wide")

# ==========================================
# INICIALIZAÇÃO DO SESSION STATE
# ==========================================
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "utilizador" not in st.session_state:
    st.session_state.utilizador = ""
if "rapv_permitido" not in st.session_state:
    st.session_state.rapv_permitido = ""
if "acesso_total" not in st.session_state:
    st.session_state.acesso_total = False
if "is_admin" not in st.session_state:
    st.session_state.is_admin = False
if "agrupar_perfil" not in st.session_state:
    st.session_state.agrupar_perfil = False
if "oficinas_agrupadas_perfil" not in st.session_state:
    st.session_state.oficinas_agrupadas_perfil = []
if "modulo_ativo" not in st.session_state:
    st.session_state.modulo_ativo = "Produtividade de Oficinas"

st.markdown("""<style>
div.stButton > button:first-child {
    font-weight: 600;
}
.modulo-header-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    margin-bottom: 4px;
}
.modulo-left {
    display: flex;
    align-items: center;
    gap: 10px;
}
.modulo-titulo {
    font-size: 18px;
    font-weight: 700;
    color: #0f172a;
    line-height: 1.2;
}
.modulo-sub {
    font-size: 12px;
    color: #475569;
    margin-bottom: 10px;
}
.modulo-badge-ativo {
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    padding: 3px 9px;
    border-radius: 999px;
    background-color: #0f766e;
    color: #ffffff;
}
.grelha-prod-wrap {
    width: 100%;
    overflow-x: auto;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    margin-top: 6px;
}
table.grelha-prod {
    width: 100%;
    border-collapse: collapse;
    font-size: 13.5px;
    color: #0f172a;
    background-color: #ffffff;
}
table.grelha-prod th {
    background-color: #f8fafc;
    color: #334155;
    font-weight: 600;
    padding: 9px 12px;
    border-bottom: 1px solid #cbd5e1;
    white-space: nowrap;
}
table.grelha-prod td {
    padding: 8px 12px;
    border-bottom: 1px solid #f1f5f9;
    white-space: nowrap;
}
table.grelha-prod tr:hover td {
    background-color: #f8fafc;
}
table.grelha-prod tr.linha-total td {
    font-weight: 700 !important;
    background-color: #eef2f6 !important;
    border-top: 2px solid #94a3b8;
    border-bottom: 2px solid #94a3b8;
    color: #0f172a;
}
.num-col {
    text-align: right;
}
.txt-col {
    text-align: left;
}
</style>""", unsafe_allow_html=True)

ICON_PRODUTIVIDADE = '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#0f766e" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="20" x2="18" y2="10"></line><line x1="12" y1="20" x2="12" y2="4"></line><line x1="6" y1="20" x2="6" y2="14"></line></svg>'

ICON_FATURACAO = '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#1d4ed8" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="5" width="20" height="14" rx="2"></rect><line x1="2" y1="10" x2="22" y2="10"></line><path d="M12 15h4"></path></svg>'

col_titulo, col_logo = st.columns([5, 1])
with col_titulo:
    st.title("Relatório de Oficinas - Grupo Motorpor")
with col_logo:
    logo_path = "logo_motorpor.jpg"
    if os.path.exists(logo_path):
        st.image(logo_path, width=150)

st.markdown("### Módulo Ativo")
col_mod1, col_mod2 = st.columns(2)

with col_mod1:
    ativo_prod = st.session_state.modulo_ativo == "Produtividade de Oficinas"
    with st.container(border=True):
        badge1 = '<span class="modulo-badge-ativo">Ativo</span>' if ativo_prod else ""
        st.markdown(f'<div class="modulo-header-row"><div class="modulo-left">{ICON_PRODUTIVIDADE}<span class="modulo-titulo">Produtividade de Oficinas</span></div>{badge1}</div><div class="modulo-sub">Tempos faturados, ausências, horas disponíveis e % produtividade</div>', unsafe_allow_html=True)
        if st.button("Produtividade de Oficinas", key="btn_mod_prod", type="primary" if ativo_prod else "secondary", use_container_width=True):
            st.session_state.modulo_ativo = "Produtividade de Oficinas"
            st.rerun()

with col_mod2:
    ativo_fat = st.session_state.modulo_ativo == "Faturação e Orçamento Diário"
    with st.container(border=True):
        badge2 = '<span class="modulo-badge-ativo">Ativo</span>' if ativo_fat else ""
        st.markdown(f'<div class="modulo-header-row"><div class="modulo-left">{ICON_FATURACAO}<span class="modulo-titulo">Faturação e Orçamento Diário</span></div>{badge2}</div><div class="modulo-sub">Faturação acumulada, orçamento diário/mensal e desvios</div>', unsafe_allow_html=True)
        if st.button("Faturação e Orçamento Diário", key="btn_mod_fat", type="primary" if ativo_fat else "secondary", use_container_width=True):
            st.session_state.modulo_ativo = "Faturação e Orçamento Diário"
            st.rerun()

st.markdown("---")
modulo_selecionado = st.session_state.modulo_ativo

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

master_file = "Gestao_Oficinas_Master.xlsx"
df_perfis = pd.DataFrame()
df_agrup_master = pd.DataFrame()
df_orc_master = pd.DataFrame()

if os.path.exists(master_file):
    try:
        xl_master = pd.ExcelFile(master_file)
        if "Perfis" in xl_master.sheet_names:
            df_perfis = xl_master.parse("Perfis")
            df_perfis["Utilizador_clean"] = df_perfis["Utilizador"].astype(str).str.strip().str.lower()
        if "Agrupamentos" in xl_master.sheet_names:
            df_agrup_master = xl_master.parse("Agrupamentos")
        if "Orçamentos" in xl_master.sheet_names:
            df_orc_master = xl_master.parse("Orçamentos")
    except Exception as e:
        st.error(f"Erro ao ler o ficheiro mestre: {e}")

tem_permissao_upload = st.session_state.get("autenticado", False) and st.session_state.get("is_admin", False)

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
    options=["Nacional (Sem Municipal)", "Beja", "Évora", "Portimão", "Faro", "Setúbal", "Sines", "Portalegre"]
)

if st.sidebar.button("Atualizar / Limpar Cache"):
    st.cache_data.clear()
    st.rerun()

ultima_atualizacao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
st.sidebar.caption(f"🔄 Última atualização: {ultima_atualizacao}")

def obter_feriados_mes(ano, mes, cidade="Nacional (Sem Municipal)"):
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
    elif "PORTALEGRE" in cidade_upper: feriados.append(date(ano, 5, 23))
    return feriados

def calcular_dias_uteis_ate_data(ano, mes, data_limite, cidade="Nacional (Sem Municipal)"):
    inicio = date(ano, mes, 1)
    fim_mes = date(ano + 1, 1, 1) - timedelta(days=1) if mes == 12 else date(ano, mes + 1, 1) - timedelta(days=1)
    limite = min(data_limite, fim_mes)
    if limite < inicio: return 0
    feriados = obter_feriados_mes(ano, mes, cidade)
    dias_uteis = 0
    atual = inicio
    while atual <= limite:
        if atual.weekday() < 5 and atual not in feriados:
            dias_uteis += 1
        atual += timedelta(days=1)
    return max(1, dias_uteis)

def calcular_dias_uteis_mes(ano, mes, cidade="Nacional (Sem Municipal)"):
    proximo_mes = date(ano + 1, 1, 1) if mes == 12 else date(ano, mes + 1, 1)
    inicio = date(ano, mes, 1)
    fim = proximo_mes - timedelta(days=1)
    feriados = obter_feriados_mes(ano, mes, cidade)
    dias_uteis = 0
    atual = inicio
    while atual <= fim:
        if atual.weekday() < 5 and atual not in feriados:
            dias_uteis += 1
        atual += timedelta(days=1)
    return dias_uteis

def parse_data_opcional(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    if not s or s.lower() in ["nan", "nat", "none", "-", ""]:
        return None
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.date()
    if isinstance(val, date):
        return val
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%y", "%d-%m-%y"):
        try:
            dt = datetime.strptime(s[:10], fmt)
            if dt.year < 2000:
                dt = dt.replace(year=dt.year + 2000)
            return dt.date()
        except Exception:
            continue
    dt_pd = pd.to_datetime(s, dayfirst=True, errors="coerce")
    if pd.notna(dt_pd):
        return dt_pd.date()
    return None

def calcular_dias_ativos_mecanico(data_entrada_val, data_saida_val, ano, mes, max_dias_decorridos, data_ref_atual, cidade="Nacional (Sem Municipal)"):
    inicio_mes = date(ano, mes, 1)
    fim_mes = date(ano + 1, 1, 1) - timedelta(days=1) if mes == 12 else date(ano, mes + 1, 1) - timedelta(days=1)

    hoje = datetime.now().date()
    if ano == hoje.year and mes == hoje.month:
        limite_periodo = min(fim_mes, data_ref_atual)
    elif date(ano, mes, 1) > hoje:
        limite_periodo = fim_mes
    else:
        limite_periodo = fim_mes

    dt_entrada = parse_data_opcional(data_entrada_val)
    dt_saida = parse_data_opcional(data_saida_val)

    inicio_efetivo = max(inicio_mes, dt_entrada) if dt_entrada is not None else inicio_mes
    fim_efetivo = min(limite_periodo, dt_saida) if dt_saida is not None else limite_periodo

    if inicio_efetivo > fim_efetivo:
        return 0

    feriados = obter_feriados_mes(ano, mes, cidade)
    dias_uteis = 0
    atual = inicio_efetivo
    while atual <= fim_efetivo:
        if atual.weekday() < 5 and atual not in feriados:
            dias_uteis += 1
        atual += timedelta(days=1)

    return min(max_dias_decorridos, dias_uteis)

dias_uteis_padrao_mes = calcular_dias_uteis_mes(ano_corrente, mes_selecionado, cidade_feriado)
data_ontem = datetime.now().date() - timedelta(days=1)
dias_decorridos_calc = calcular_dias_uteis_ate_data(ano_corrente, mes_selecionado, data_ontem, cidade_feriado)
decorridos_iniciais = dias_uteis_padrao_mes if (mes_selecionado < mes_atual_num or ano_corrente < 2026) else min(dias_decorridos_calc, dias_uteis_padrao_mes)

dias_uteis_padrao_mes = st.sidebar.slider(
    f"Total de Dias Úteis ({meses_pt[mes_selecionado].capitalize()}):", 
    min_value=15, max_value=25, value=dias_uteis_padrao_mes
)

dias_decorridos = st.sidebar.number_input(
    "Dias úteis decorridos:", 
    min_value=1, max_value=dias_uteis_padrao_mes, 
    value=decorridos_iniciais if (mes_selecionado == mes_atual_num and ano_corrente == 2026) else (dias_uteis_padrao_mes if mes_selecionado <= mes_atual_num else 1)
)

percentagem_progresso = (dias_decorridos / dias_uteis_padrao_mes) * 100 if dias_uteis_padrao_mes > 0 else 0
progresso_decimal = min(max(dias_decorridos / dias_uteis_padrao_mes, 0.0), 1.0)

st.sidebar.progress(
    progresso_decimal,
    text=f"Progresso: {dias_decorridos} de {dias_uteis_padrao_mes} dias ({percentagem_progresso:.1f}%)"
)

if tem_permissao_upload:
    st.sidebar.markdown("---")
    st.sidebar.header("Atualizar Ficheiros (Admin opcional)")
    st.sidebar.file_uploader("Substituir 'tabela mecânicos'", type=["xlsx"], key="up_mec")
    st.sidebar.file_uploader("Substituir 'Comparativo de Horas'", type=["xlsx"], key="up_hf")
    st.sidebar.file_uploader("Substituir 'Listagem de Ausencias'", type=["xlsx"], key="up_aus")
    st.sidebar.file_uploader("Mapa de Faturação 1 (CSV)", type=["csv"], key="up_f1")
    st.sidebar.file_uploader("Mapa de Faturação 2 (CSV)", type=["csv"], key="up_f2")

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
                    match = df_perfis[df_perfis["Utilizador_clean"] == user_input]
                    if not match.empty:
                        st.session_state.autenticado = True
                        st.session_state.utilizador = match.iloc[0]["Utilizador"]
                        rapv_perm = str(match.iloc[0]["RAPV Permitidos"]).strip()
                        
                        st.session_state.rapv_permitido = rapv_perm
                        st.session_state.is_admin = (rapv_perm.lower() == "administrador")
                        st.session_state.acesso_total = (rapv_perm.lower() in ["administrador", "todas"])
                        
                        agrupar_val = str(match.iloc[0].get("Agrupar", "não")).strip().lower()
                        st.session_state.agrupar_perfil = (agrupar_val == "sim")
                        
                        agrupamento_str = str(match.iloc[0].get("Agrupamento", "")).strip()
                        if agrupamento_str and agrupamento_str.lower() != "nan" and agrupamento_str.lower() != "não":
                            st.session_state.oficinas_agrupadas_perfil = [o.strip().upper() for o in agrupamento_str.split(",") if o.strip()]
                        else:
                            st.session_state.oficinas_agrupadas_perfil = []
                        st.rerun()
                    else:
                        st.error("Utilizador ou password incorretos.")
                else:
                    st.warning("Preencha o utilizador e a password.")
        else:
            st.error("Ficheiro 'Gestao_Oficinas_Master.xlsx' sem aba de perfis.")
    else:
        st.success(f"Sessão iniciada como: **{st.session_state.utilizador}**")
        if st.button("Logout / Mudar Utilizador"):
            st.session_state.autenticado = False
            st.session_state.utilizador = ""
            st.session_state.rapv_permitido = ""
            st.session_state.acesso_total = False
            st.session_state.is_admin = False
            st.session_state.agrupar_perfil = False
            st.session_state.oficinas_agrupadas_perfil = []
            st.rerun()

file_mecanicos_up = st.session_state.get("up_mec")
file_hf_up = st.session_state.get("up_hf")
file_ausencias_up = st.session_state.get("up_aus")
file_fact1_up = st.session_state.get("up_f1")
file_fact2_up = st.session_state.get("up_f2")

path_mecanicos = file_mecanicos_up if file_mecanicos_up is not None else ("tabela mecânicos.xlsx" if os.path.exists("tabela mecânicos.xlsx") else None)

nome_arq_hf = f"Comparativo de Horas_{sigla_mes}_{ano_corrente}.xlsx"
path_hf = file_hf_up if file_hf_up is not None else (nome_arq_hf if os.path.exists(nome_arq_hf) else None)

nome_arq_aus = f"Listagem de AusenciasPresenças_{sigla_mes}_{ano_corrente}.xlsx"
path_ausencias = file_ausencias_up if file_ausencias_up is not None else (nome_arq_aus if os.path.exists(nome_arq_aus) else None)

def converter_valor_pt(val_str):
    if pd.isna(val_str): return 0.0
    s = str(val_str).replace("€", "").strip()
    if not s: return 0.0
    s = s.replace("−", "-").replace("–", "-")
    if "." in s and "," in s:
        if s.rfind('.') > s.rfind(','):
            s = s.replace(",", "")
        else:
            s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try: return float(s)
    except: return 0.0

def converter_data_robusta(val_data, ano_alvo, mes_alvo):
    if pd.isna(val_data):
        return datetime(ano_alvo, mes_alvo, 1)
    
    s = str(val_data).strip()
    if not s or s.lower() == "nan":
        return datetime(ano_alvo, mes_alvo, 1)

    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%y", "%d-%m-%y"):
        try:
            dt = datetime.strptime(s[:10], fmt)
            if dt.year < 2000:
                dt = dt.replace(year=dt.year + 2000)
            return dt
        except:
            continue
            
    dt_pd = pd.to_datetime(s, dayfirst=True, errors="coerce")
    if pd.notna(dt_pd):
        return dt_pd.to_pydatetime()
        
    dt_pd2 = pd.to_datetime(s, dayfirst=False, errors="coerce")
    if pd.notna(dt_pd2):
        return dt_pd2.to_pydatetime()
        
    return datetime(ano_alvo, mes_alvo, 1)

def limpar_str(s):
    if pd.isna(s): return ""
    s_nfkd = unicodedata.normalize('NFKD', str(s))
    return "".join([c for c in s_nfkd if not unicodedata.combining(c)]).upper().strip().rstrip(".")

def limpar_prefixo_oficina(nome):
    limpo = str(nome).strip()
    if limpo.upper().startswith("OF. "): limpo = limpo[4:].strip()
    elif limpo.upper().startswith("OF "): limpo = limpo[3:].strip()
    return limpo.rstrip(".")

def oficina_permitida_check(oficina_nome):
    if st.session_state.acesso_total: return True
    oficina_up = limpar_str(oficina_nome)
    if oficina_up == "SEM PRODUTIVO": return True
    
    for permitida in st.session_state.oficinas_agrupadas_perfil:
        p_clean = limpar_str(permitida)
        p_clean_sem_of = limpar_prefixo_oficina(p_clean)
        oficina_sem_of = limpar_prefixo_oficina(oficina_up)
        if (p_clean == oficina_up or 
            p_clean_sem_of == oficina_sem_of or 
            p_clean in oficina_up or 
            oficina_up in p_clean or
            p_clean_sem_of in oficina_sem_of or
            oficina_sem_of in p_clean_sem_of):
            return True
    return False

def carregar_dados_csv_faturacao():
    df_fact_total = pd.DataFrame()
    csv_fontes = [file_fact1_up, file_fact2_up]
    
    if all(f is None for f in csv_fontes):
        encontrados = []
        dir_base = os.path.dirname(os.path.abspath(__file__)) if "__file__" in locals() else "."
        
        for d in [dir_base, "."]:
            for ext in ["*.csv", "*.CSV"]:
                encontrados.extend(glob.glob(os.path.join(d, ext)))
        
        fich_exatos = [
            "MOTOREX_MAPA_FACTURAÇÃO__2026_11.csv",
            "MULTIAUTO_MAPA_FACTURAÇÃO__2026_11.csv"
        ]
        for fe in fich_exatos:
            caminho_fe = os.path.join(dir_base, fe)
            if os.path.exists(caminho_fe):
                encontrados.append(caminho_fe)
            elif os.path.exists(fe):
                encontrados.append(fe)

        encontrados = list(set(encontrados))
        
        for arq in encontrados:
            csv_fontes.append(arq)

    for file_f in csv_fontes:
        if file_f is not None:
            try:
                if hasattr(file_f, "getvalue"):
                    conteudo = file_f.getvalue().decode("iso-8859-1", errors="ignore")
                elif hasattr(file_f, "read"):
                    file_f.seek(0)
                    res = file_f.read()
                    conteudo = res.decode("iso-8859-1", errors="ignore") if isinstance(res, bytes) else res
                else:
                    with open(file_f, "r", encoding="iso-8859-1", errors="ignore") as f:
                        conteudo = f.read()
                
                linhas_validas = []
                for sep in [";", "\t", ","]:
                    f_io = io.StringIO(conteudo)
                    reader = csv.reader(f_io, delimiter=sep)
                    temp_linhas = [row for row in reader if len(row) > 2]
                    if len(temp_linhas) > 5:
                        linhas_validas = temp_linhas
                        break
                
                if not linhas_validas:
                    f_io = io.StringIO(conteudo)
                    reader = csv.reader(f_io, delimiter=";")
                    linhas_validas = [row for row in reader if len(row) > 0]

                if linhas_validas:
                    df_temp = pd.DataFrame(linhas_validas)
                    df_temp["Origem_CSV"] = getattr(file_f, "name", str(file_f))
                    df_fact_total = pd.concat([df_fact_total, df_temp], ignore_index=True)
            except Exception as e:
                st.error(f"Erro ao ler CSV: {e}")
    return df_fact_total

# ==========================================
# CORPO PRINCIPAL
# ==========================================
if not st.session_state.autenticado:
    st.warning("🔒 Por favor, faça login na barra lateral para aceder aos relatórios.")
elif not path_mecanicos or not path_hf:
    st.warning(f"⚠️ Não foram encontrados os ficheiros padrão no disco para o período de {meses_pt[mes_selecionado].capitalize()} {ano_corrente} ({sigla_mes}). Certifique-se de que os ficheiros estão na pasta do projeto.")
else:
    if modulo_selecionado == "Produtividade de Oficinas":
        xl = pd.ExcelFile(path_mecanicos)
        df_mecanicos = xl.parse("Mecânicos" if "Mecânicos" in xl.sheet_names else xl.sheet_names[0])
        df_hf = pd.read_excel(path_hf)
        
        df_ausencias = pd.DataFrame()
        if path_ausencias is not None:
            try:
                df_ausencias = pd.read_excel(path_ausencias)
            except:
                pass

        df_ausencias_grouped = pd.DataFrame(columns=["Técnico_clean", "Ausencias_calc"])
        if not df_ausencias.empty:
            col_data_aus = next((c for c in ["Data", "Data (Data)", "Dia"] if c in df_ausencias.columns), df_ausencias.columns[4] if len(df_ausencias.columns) > 4 else None)
            col_desc_aus = df_ausencias.columns[1] if len(df_ausencias.columns) > 1 else None
            col_op_aus = next((c for c in ["Mecânico (Nome)", "Nome", "Técnico", "Tecnico", "Operário", "Operario"] if c in df_ausencias.columns), df_ausencias.columns[3] if len(df_ausencias.columns) > 3 else df_ausencias.select_dtypes(include=['object']).columns[0])
            col_val_aus = df_ausencias.select_dtypes(include=['number']).columns[-1]

            if col_desc_aus and col_val_aus:
                df_aus_filtrado = df_ausencias.copy()
                if col_data_aus:
                    df_aus_filtrado["Data_Parsed"] = pd.to_datetime(df_aus_filtrado[col_data_aus], dayfirst=True, errors="coerce")
                    if ano_corrente == datetime.now().year and mes_selecionado == datetime.now().month:
                        df_aus_filtrado = df_aus_filtrado[df_aus_filtrado["Data_Parsed"].dt.date <= data_ontem]

                df_aus_filtrado = df_aus_filtrado[df_aus_filtrado[col_desc_aus].astype(str).str.strip().str.lower() != "presença"]
                df_aus_filtrado["Técnico_clean"] = df_aus_filtrado[col_op_aus].astype(str).str.strip().str.upper()
                df_ausencias_grouped = df_aus_filtrado.groupby("Técnico_clean", as_index=False)[col_val_aus].sum()
                df_ausencias_grouped.rename(columns={col_val_aus: "Ausencias_calc"}, inplace=True)

        df_mecanicos.rename(columns={"Técnicos": "Técnico", "Nome": "Técnico"}, inplace=True)
        df_mecanicos["Técnico_clean"] = df_mecanicos["Técnico"].astype(str).str.strip().str.upper()
        df_mecanicos["Oficina"] = df_mecanicos["Oficina"].astype(str).str.strip()
        df_mecanicos["Local"] = df_mecanicos["Local"].astype(str).str.strip()

        col_entrada = next((c for c in df_mecanicos.columns if limpar_str(c) in ["DATA ENTRADA", "DATA DE ENTRADA", "ENTRADA"]), None)
        col_saida = next((c for c in df_mecanicos.columns if limpar_str(c) in ["DATA SAIDA", "DATA DE SAIDA", "SAIDA"]), None)

        df_mecanicos["Dias_Ativos_Calc"] = df_mecanicos.apply(
            lambda r: calcular_dias_ativos_mecanico(
                r[col_entrada] if col_entrada else None,
                r[col_saida] if col_saida else None,
                ano_corrente,
                mes_selecionado,
                dias_decorridos,
                data_ontem,
                cidade_feriado
            ),
            axis=1
        )

        mapa_prod_master = {}
        if not df_agrup_master.empty:
            for _, r in df_agrup_master.iterrows():
                ofic_orig = str(r.iloc[0]).strip()
                agrup_prod = str(r.iloc[5]).strip() if len(r) > 5 and pd.notna(r.iloc[5]) else ofic_orig
                if agrup_prod and agrup_prod.lower() != "nan":
                    mapa_prod_master[limpar_str(ofic_orig)] = agrup_prod

        def aplicar_agrupamento_prod(nome_ofic):
            u = limpar_str(nome_ofic)
            return mapa_prod_master.get(u, nome_ofic)

        df_mecanicos["Oficina_Original"] = df_mecanicos["Oficina"]
        df_mecanicos["Oficina"] = df_mecanicos["Oficina"].apply(aplicar_agrupamento_prod)

        oficinas_oficiais = df_mecanicos["Oficina"].unique().tolist()
        dict_oficina_local = dict(zip(df_mecanicos["Oficina"].str.upper(), df_mecanicos["Local"]))

        def normalizar_oficina_e_local(nome_oficina):
            if pd.isna(nome_oficina): return "Geral", "Sem produtivo"
            nome_str = str(nome_oficina).strip()
            agrupado = aplicar_agrupamento_prod(nome_str)
            for ofic_oficial in oficinas_oficiais:
                if ofic_oficial.upper() == agrupado.upper():
                    return ofic_oficial, dict_oficina_local.get(ofic_oficial.upper(), "Geral")
            return agrupado, "Sem produtivo"

        candidatos_op = ["Operário", "Operario", "Técnico", "Tecnico", "Nome", "Mecânico", "Mecanico"]
        col_operario_hf = next((c for c in candidatos_op if c in df_hf.columns), df_hf.select_dtypes(include=['object']).columns[0])
        col_fact = next((c for c in ["Tempos facturados", "Tempos facturados ", "Horas facturadas", "Facturado", "Valor"] if c in df_hf.columns), df_hf.select_dtypes(include=['number']).columns[-1])
        col_oficina_hf = next((c for c in ["Oficina", "Escritório", "Escritorio", "Local"] if c in df_hf.columns), None)
        
        df_hf["Operario_clean"] = df_hf[col_operario_hf].astype(str).str.strip().str.upper()
        
        if col_oficina_hf:
            res_norm = df_hf[col_oficina_hf].apply(normalizar_oficina_e_local)
            df_hf["Oficina_Normalizada"] = [r[0] for r in res_norm]
            df_hf["Local_Normalizado"] = [r[1] for r in res_norm]
        else:
            df_hf["Oficina_Normalizada"] = "Geral"
            df_hf["Local_Normalizado"] = "Sem produtivo"

        df_hf_known = df_hf[df_hf["Operario_clean"].notna() & (df_hf["Operario_clean"] != "") & (df_hf["Operario_clean"] != "NAN")].copy()
        df_hf_sem_op = df_hf[df_hf["Operario_clean"].isna() | (df_hf["Operario_clean"] == "") | (df_hf["Operario_clean"] == "NAN")].copy()

        df_hf_grouped = df_hf_known.groupby("Operario_clean", as_index=False)[col_fact].sum()
        df_hf_grouped.rename(columns={col_fact: "Horas facturadas", "Operario_clean": "Técnico_clean"}, inplace=True)

        df_merged = pd.merge(df_mecanicos, df_hf_grouped, on="Técnico_clean", how="left")
        if not df_ausencias_grouped.empty:
            df_merged = pd.merge(df_merged, df_ausencias_grouped, on="Técnico_clean", how="left")
            df_merged["Ausencias"] = df_merged["Ausencias_calc"].fillna(0.0)
        else:
            df_merged["Ausencias"] = 0.0

        df_merged["Horas facturadas"] = df_merged["Horas facturadas"].fillna(0.0)
        df_merged.loc[(df_merged["Dias_Ativos_Calc"] <= 0) & (df_merged["Horas facturadas"] == 0), "Ausencias"] = 0.0
        df_merged["Horas potencial dia"] = df_merged["Dias_Ativos_Calc"] * df_merged["Horas reais dia"]
        df_merged["Horas disponiveis"] = df_merged["Horas potencial dia"] - df_merged["Ausencias"]

        df_merged = df_merged[
            (df_merged["Dias_Ativos_Calc"] > 0) |
            (df_merged["Horas facturadas"] != 0)
        ].copy()

        colunas_finais = ["Técnico", "Local", "Oficina", "Função", "Horas potencial dia", "Horas reais dia", "Ausencias", "Horas disponiveis", "Horas facturadas", "Produtividade %"]
        
        df_merged["Produtividade %"] = 0.0
        mask_disp = df_merged["Horas disponiveis"] > 0
        df_merged.loc[mask_disp, "Produtividade %"] = (df_merged.loc[mask_disp, "Horas facturadas"] / df_merged.loc[mask_disp, "Horas disponiveis"]) * 100

        df_exibicao_tecnicos = df_merged[[c for c in colunas_finais if c in df_merged.columns]].copy()

        linhas_sem_produtivo = []
        if not df_hf_sem_op.empty:
            df_sem_op_grouped = df_hf_sem_op.groupby(["Local_Normalizado", "Oficina_Normalizada"], as_index=False)[col_fact].sum()
            for _, row in df_sem_op_grouped.iterrows():
                if row[col_fact] > 0:
                    linhas_sem_produtivo.append({
                        "Técnico": "Sem produtivo associado", "Local": row["Local_Normalizado"], "Oficina": row["Oficina_Normalizada"],
                        "Função": "Sem Produtivo", "Horas potencial dia": 0.0, "Horas reais dia": 0.0, "Ausencias": 0.0,
                        "Horas disponiveis": 0.0, "Horas facturadas": row[col_fact], "Produtividade %": 0.0
                    })

        df_sem_prod_final = pd.DataFrame(linhas_sem_produtivo)
        df_completo = pd.concat([df_exibicao_tecnicos, df_sem_prod_final], ignore_index=True) if not df_sem_prod_final.empty else df_exibicao_tecnicos

        def checar_perm_prod(oficina_nome):
            if st.session_state.acesso_total: return True
            if oficina_permitida_check(oficina_nome): return True
            
            oficina_up = limpar_str(oficina_nome)
            user_rapv = limpar_str(st.session_state.rapv_permitido)
            if not df_agrup_master.empty:
                for _, r in df_agrup_master.iterrows():
                    agrup_p = limpar_str(r.iloc[5] if len(r) > 5 and pd.notna(r.iloc[5]) else r.iloc[0])
                    rapv_m = limpar_str(r.iloc[4] if len(r) > 4 and pd.notna(r.iloc[4]) else "")
                    if rapv_m == user_rapv and (agrup_p == oficina_up or oficina_up in agrup_p or agrup_p in oficina_up or limpar_prefixo_oficina(agrup_p) in limpar_prefixo_oficina(oficina_up)):
                        return True
            return False

        df_completo = df_completo[df_completo["Oficina"].apply(checar_perm_prod)]

        st.subheader(f"Mapa de Produtividade - {meses_pt[mes_selecionado].capitalize()} {ano_corrente}")

        col_f_local, col_f_oficina = st.columns(2)
        with col_f_local:
            locais_unicos = sorted(df_completo["Local"].dropna().astype(str).unique().tolist())
            locais_disponiveis = (["Todas"] + locais_unicos) if len(locais_unicos) > 1 else locais_unicos
            sel_local = st.selectbox("Selecionar Cidade (Local)", locais_disponiveis if locais_disponiveis else ["Sem acesso"])

        df_filtrado_local = df_completo[df_completo["Local"] == sel_local] if sel_local != "Todas" else df_completo

        with col_f_oficina:
            oficinas_unicas = sorted(df_filtrado_local["Oficina"].dropna().astype(str).unique().tolist())
            oficinas_disponiveis = (["Todas"] + oficinas_unicas) if len(oficinas_unicas) > 1 else oficinas_unicas
            sel_oficina = st.selectbox("Selecionar Oficina", oficinas_disponiveis if oficinas_disponiveis else ["Sem acesso"])

        df_final_filtro = df_filtrado_local[df_filtrado_local["Oficina"] == sel_oficina] if sel_oficina != "Todas" else df_filtrado_local

        tot_h_potencial = df_final_filtro["Horas potencial dia"].sum()
        tot_h_reais = df_final_filtro["Horas reais dia"].sum()
        tot_ausencias = df_final_filtro["Ausencias"].sum()
        tot_h_disp = df_final_filtro["Horas disponiveis"].sum()
        tot_h_fact = df_final_filtro["Horas facturadas"].sum()
        tot_prod = (tot_h_fact / tot_h_disp * 100) if tot_h_disp > 0 else 0.0

        cols_visiveis = [c for c in colunas_finais if c in df_final_filtro.columns]
        ths_html = "".join([
            f'<th class="{"num-col" if c in ["Horas potencial dia", "Horas reais dia", "Ausencias", "Horas disponiveis", "Horas facturadas", "Produtividade %"] else "txt-col"}">{c}</th>'
            for c in cols_visiveis
        ])

        trs_html = []
        for _, r in df_final_filtro.iterrows():
            tds = []
            for c in cols_visiveis:
                val = r[c]
                if c in ["Horas potencial dia", "Horas reais dia", "Ausencias", "Horas disponiveis"]:
                    tds.append(f'<td class="num-col">{float(val):.1f}</td>')
                elif c == "Horas facturadas":
                    tds.append(f'<td class="num-col">{float(val):.2f}</td>')
                elif c == "Produtividade %":
                    tds.append(f'<td class="num-col">{float(val):.1f}%</td>')
                else:
                    tds.append(f'<td class="txt-col">{val}</td>')
            trs_html.append(f'<tr>{"".join(tds)}</tr>')

        totais_map = {
            "Técnico": "<strong>TOTAL</strong>",
            "Local": "-",
            "Oficina": "-",
            "Função": "-",
            "Horas potencial dia": f"<strong>{tot_h_potencial:.1f}</strong>",
            "Horas reais dia": f"<strong>{tot_h_reais:.1f}</strong>",
            "Ausencias": f"<strong>{tot_ausencias:.1f}</strong>",
            "Horas disponiveis": f"<strong>{tot_h_disp:.1f}</strong>",
            "Horas facturadas": f"<strong>{tot_h_fact:.2f}</strong>",
            "Produtividade %": f"<strong>{tot_prod:.1f}%</strong>"
        }
        tds_total = "".join([
            f'<td class="{"num-col" if c in ["Horas potencial dia", "Horas reais dia", "Ausencias", "Horas disponiveis", "Horas facturadas", "Produtividade %"] else "txt-col"}">{totais_map.get(c, "-")}</td>'
            for c in cols_visiveis
        ])
        trs_html.append(f'<tr class="linha-total">{tds_total}</tr>')

        tabela_html = f'<div class="grelha-prod-wrap"><table class="grelha-prod"><thead><tr>{ths_html}</tr></thead><tbody>{"".join(trs_html)}</tbody></table></div>'
        st.markdown(tabela_html, unsafe_allow_html=True)

    elif modulo_selecionado == "Faturação e Orçamento Diário":
        st.subheader(f"Faturação e Orçamento Diário - {meses_pt[mes_selecionado].capitalize()} {ano_corrente}")
        
        df_orc = df_orc_master.copy()
        df_fact_total = carregar_dados_csv_faturacao()
                
        if df_fact_total.empty:
            st.warning("⚠️ Não foram encontrados ficheiros CSV de faturação no disco.")
        else:
            num_cols = df_fact_total.shape[1]
            idx_ofic = 0
            idx_val = 24 if num_cols > 24 else num_cols - 2
            idx_data = 28 if num_cols > 28 else num_cols - 1

            df_fact_total["Oficina_CSV"] = df_fact_total.iloc[:, idx_ofic].astype(str).str.strip()
            invalidos_csv = ["NAN", "", "TOTAL", "MAPA DE FACTURAÇÃO", "OFICINA", "OFICINAS", "NOME", "SECCAO"]
            df_fact_total = df_fact_total[~df_fact_total["Oficina_CSV"].str.upper().isin(invalidos_csv)]
            df_fact_total = df_fact_total[~df_fact_total["Oficina_CSV"].str.startswith("*")]

            df_fact_total["Data_Parsed"] = df_fact_total.iloc[:, idx_data].apply(lambda x: converter_data_robusta(x, ano_corrente, mes_selecionado))
            df_fact_total["Valor_Num"] = df_fact_total.iloc[:, idx_val].apply(converter_valor_pt)

            mapa_fact_master = {}
            mapa_rapv_master = {}
            lista_suboficinas_por_grupo_fact = {}

            if not df_agrup_master.empty:
                for _, r in df_agrup_master.iterrows():
                    ofic_orig = str(r.iloc[0]).strip()
                    agrup_fact = str(r.iloc[6]).strip() if len(r) > 6 and pd.notna(r.iloc[6]) else ofic_orig
                    rapv_val = str(r.iloc[4]).strip() if len(r) > 4 and pd.notna(r.iloc[4]) else "Geral"
                    
                    if not agrup_fact or agrup_fact.lower() == "nan":
                        agrup_fact = ofic_orig

                    ofic_clean = limpar_str(ofic_orig)
                    mapa_fact_master[ofic_clean] = agrup_fact
                    mapa_rapv_master[agrup_fact] = rapv_val

                    if agrup_fact not in lista_suboficinas_por_grupo_fact:
                        lista_suboficinas_por_grupo_fact[agrup_fact] = set()
                    lista_suboficinas_por_grupo_fact[agrup_fact].add(ofic_clean)

            def obter_grupo_fact(nome_csv):
                u = limpar_str(nome_csv)
                if u in mapa_fact_master:
                    return mapa_fact_master[u]
                
                is_colisao = "COLISÃO" in u or "COLISAO" in u
                melhor_match = None
                max_tokens = 0
                
                for chave_master, agrup in mapa_fact_master.items():
                    agrup_is_colisao = "COLISÃO" in agrup.upper() or "COLISAO" in agrup.upper()
                    if is_colisao != agrup_is_colisao:
                        continue
                        
                    tokens_master = [t for t in chave_master.replace("OF.", "").replace("COLISÃO", "").replace("COLISAO", "").split() if t not in ["DE", "DA", "E"]]
                    if tokens_master and all(t in u for t in tokens_master):
                        if len(tokens_master) > max_tokens:
                            max_tokens = len(tokens_master)
                            melhor_match = agrup
                            
                if melhor_match:
                    return melhor_match
                    
                for agrup in mapa_fact_master.values():
                    if limpar_str(agrup) in u or u in limpar_str(agrup):
                        return agrup
                        
                return nome_csv

            df_fact_total["Oficina_Agrupada"] = df_fact_total["Oficina_CSV"].apply(obter_grupo_fact)
            df_fact_total["RAPV"] = df_fact_total["Oficina_Agrupada"].map(lambda g: mapa_rapv_master.get(g, "Geral"))

            oficinas_disponiveis_fat = sorted(list(set(df_fact_total["Oficina_Agrupada"].dropna().astype(str).unique())))

            if st.session_state.acesso_total:
                oficinas_autorizadas_lista = oficinas_disponiveis_fat
            else:
                user_rapv = str(st.session_state.rapv_permitido).strip().upper()
                oficinas_autorizadas_lista = [
                    o for o in oficinas_disponiveis_fat 
                    if mapa_rapv_master.get(o, "").strip().upper() == user_rapv or oficina_permitida_check(o)
                ]
            if not oficinas_autorizadas_lista: oficinas_autorizadas_lista = ["Sem acesso"]

            col_sel_rapv, col_sel_ofic = st.columns(2)
            
            if st.session_state.acesso_total:
                rapvs_autorizados = sorted(list(set(mapa_rapv_master.values())))
            else:
                rapv_usuario_atual = str(st.session_state.rapv_permitido).strip()
                rapvs_autorizados = [rapv_usuario_atual] if rapv_usuario_atual else ["Geral"]

            with col_sel_rapv:
                rapv_selecionado = st.selectbox("Selecionar RAPV:", options=rapvs_autorizados)

            oficinas_filtradas_rapv = [
                o for o in oficinas_autorizadas_lista 
                if mapa_rapv_master.get(o, "").strip().upper() == rapv_selecionado.strip().upper()
            ]
            if not oficinas_filtradas_rapv: oficinas_filtradas_rapv = [oficinas_autorizadas_lista[0]]

            with col_sel_ofic:
                oficina_selecionada_fat = st.selectbox("Selecionar Oficina (Faturação):", options=oficinas_filtradas_rapv)
                
            sel_clean = limpar_str(oficina_selecionada_fat)
            suboficinas_grupo = lista_suboficinas_por_grupo_fact.get(oficina_selecionada_fat, {sel_clean})
            
            df_oficina_fat = df_fact_total[
                (df_fact_total["Oficina_Agrupada"].apply(limpar_str) == sel_clean) |
                (df_fact_total["Oficina_CSV"].apply(limpar_str).isin(suboficinas_grupo))
            ]
            
            df_oficina_fat = df_oficina_fat[(df_oficina_fat["Data_Parsed"].dt.month == mes_selecionado) & (df_oficina_fat["Data_Parsed"].dt.year == ano_corrente)]
            
            df_diario = df_oficina_fat.groupby(df_oficina_fat["Data_Parsed"].dt.date, as_index=False)["Valor_Num"].sum()
            df_diario.rename(columns={"Valor_Num": "Facturacao_Diaria", "Data_Parsed": "Data"}, inplace=True)
            df_diario = df_diario.dropna(subset=["Data"]).sort_values("Data")
            
            inicio_mes = date(ano_corrente, mes_selecionado, 1)
            fim_mes = date(ano_corrente + 1, 1, 1) - timedelta(days=1) if mes_selecionado == 12 else date(ano_corrente, mes_selecionado + 1, 1) - timedelta(days=1)
            feriados_mes = obter_feriados_mes(ano_corrente, mes_selecionado, cidade_feriado)
            
            dias_uteis_lista = []
            atual = inicio_mes
            while atual <= fim_mes:
                if atual.weekday() < 5 and atual not in feriados_mes: dias_uteis_lista.append(atual)
                atual += timedelta(days=1)
                
            df_completo_dias = pd.DataFrame({"Data": dias_uteis_lista})
            df_completo_dias = pd.merge(df_completo_dias, df_diario, on="Data", how="left").fillna(0.0)
            df_completo_dias["Data_Str"] = pd.to_datetime(df_completo_dias["Data"]).dt.strftime("%d/%m/%Y")
            
            hoje_atual = datetime.now().date()
            data_limite_grafico = hoje_atual if (ano_corrente == hoje_atual.year and mes_selecionado == hoje_atual.month) else fim_mes

            df_completo_dias["Facturacao_Acumulada"] = df_completo_dias["Facturacao_Diaria"].cumsum()
            df_completo_dias["Facturacao_Acumulada_Graf"] = df_completo_dias.apply(lambda r: r["Facturacao_Acumulada"] if r["Data"] <= data_limite_grafico else None, axis=1)
            df_completo_dias["Facturacao_Diaria_Graf"] = df_completo_dias.apply(lambda r: r["Facturacao_Diaria"] if r["Data"] <= data_limite_grafico else 0.0, axis=1)

            orcamento_mensal = 0.0
            if not df_orc.empty:
                col_ofic = df_orc.columns[0]
                col_mes = df_orc.columns[3]
                col_val = df_orc.columns[4]
                nome_mes_alvo = limpar_str(meses_pt[mes_selecionado])
                
                for _, row in df_orc.iterrows():
                    ofic_raw = limpar_prefixo_oficina(str(row[col_ofic]))
                    ofic_row = limpar_str(ofic_raw)
                    if not ofic_row or any(exc in ofic_row for exc in ["TOTAL", "GRUPO", "MAPA", "ACUMULADO"]): continue
                    
                    val_mes_cell = row[col_mes]
                    mes_encontrado = False
                    if isinstance(val_mes_cell, (pd.Timestamp, datetime, date)):
                        if val_mes_cell.month == mes_selecionado: mes_encontrado = True
                    else:
                        mes_row_str = limpar_str(val_mes_cell)
                        if nome_mes_alvo in mes_row_str or mes_row_str in nome_mes_alvo: mes_encontrado = True
                        else:
                            try:
                                if int(val_mes_cell) == mes_selecionado: mes_encontrado = True
                            except: pass
                            
                    if mes_encontrado:
                        if sel_clean == ofic_row or sel_clean in ofic_row or ofic_row in sel_clean:
                            orcamento_mensal += converter_valor_pt(row[col_val])
                                
            orcamento_diario = orcamento_mensal / dias_uteis_padrao_mes if dias_uteis_padrao_mes > 0 else 0.0
            
            orcamentos_ac = []
            acum = 0.0
            for d in df_completo_dias["Data"]:
                acum += orcamento_diario
                orcamentos_ac.append(acum)
            df_completo_dias["Orcamento_Acumulado"] = orcamentos_ac
            
            val_fat_reais = df_diario[df_diario["Data"] <= data_limite_grafico]["Facturacao_Diaria"]
            facturacao_a_data = val_fat_reais.sum() if not val_fat_reais.empty else 0.0
            
            dias_uteis_decorridos_reais = len(df_completo_dias[df_completo_dias["Data"] <= data_limite_grafico])
            orcamento_a_data = dias_uteis_decorridos_reais * orcamento_diario
            desvio_orcamento = facturacao_a_data - orcamento_a_data
            
            col_graf, col_kpi = st.columns([3, 1])
            
            with col_graf:
                base = alt.Chart(df_completo_dias).encode(
                    x=alt.X('Data_Str:N', title='Data', sort=list(df_completo_dias["Data_Str"]), axis=alt.Axis(labelAngle=-45))
                )
                barras = base.mark_bar(color='#115e59').encode(
                    y=alt.Y('Facturacao_Acumulada_Graf:Q', title='Valor (€)'),
                    tooltip=[
                        alt.Tooltip('Data_Str:N', title='Data'),
                        alt.Tooltip('Facturacao_Diaria_Graf:Q', title='Faturação Diária', format=',.2f'),
                        alt.Tooltip('Facturacao_Acumulada_Graf:Q', title='Faturação Acumulada', format=',.2f')
                    ]
                )
                linha_orc = base.mark_line(color='#f97316', strokeWidth=3).encode(
                    y=alt.Y('Orcamento_Acumulado:Q', title='Orçamento Acumulado'),
                    tooltip=[
                        alt.Tooltip('Data_Str:N', title='Data'),
                        alt.Tooltip('Orcamento_Acumulado:Q', title='Orçamento Acumulado', format=',.2f')
                    ]
                )
                chart = (barras + linha_orc).properties(height=420).interactive()
                st.altair_chart(chart, use_container_width=True)
                
            with col_kpi:
                orc_mensal_txt = f"{orcamento_mensal:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
                orc_diario_txt = f"{orcamento_diario:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
                fat_data_txt = f"{facturacao_a_data:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
                desvio_txt = f"{desvio_orcamento:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
                cor_desvio = "green" if desvio_orcamento >= 0 else "red"

                st.metric(label="ORÇAMENTO MENSAL", value=orc_mensal_txt)
                st.metric(label="ORÇAMENTO DIÁRIO", value=orc_diario_txt)
                st.metric(label="FATURAÇÃO À DATA", value=fat_data_txt)
                
                st.markdown(f'<div style="font-size: 14px; font-weight: 600; color: rgb(49, 51, 63); margin-bottom: 0px;">DESVIO ORÇAMENTO</div><div style="font-size: 24px; font-weight: 700; color: {cor_desvio}; margin-top: 0px; margin-bottom: 0px;">{desvio_txt}</div>', unsafe_allow_html=True)
                st.metric(label="", value="", delta=desvio_txt, label_visibility="collapsed")