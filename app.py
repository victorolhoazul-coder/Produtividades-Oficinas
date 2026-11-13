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
    st.session_state.agrupar_perfil = False
    st.session_state.oficinas_agrupadas_perfil = []

df_perfis = pd.DataFrame()
if os.path.exists("perfis_utilizadores.xlsx"):
    try:
        xl_perfis = pd.ExcelFile("perfis_utilizadores.xlsx")
        sheet_perfis_nome = "Agrupamento_produtividade" if "Agrupamento_produtividade" in xl_perfis.sheet_names else xl_perfis.sheet_names[0]
        df_perfis = xl_perfis.parse(sheet_perfis_nome)
        df_perfis["Utilizador_clean"] = df_perfis["Utilizador"].astype(str).str.strip().str.lower()
    except Exception as e:
        st.error(f"Erro ao ler perfis_utilizadores.xlsx: {e}")

utilizadores_com_permissao_upload = ["vazul", "mchegadinho"]
utilizador_atual = st.session_state.utilizador.strip().lower() if st.session_state.autenticado else ""
tem_permissao_upload = st.session_state.autenticado and (utilizador_atual in [u.lower() for u in utilizadores_com_permissao_upload])

st.sidebar.header("Navegação e Módulo")
modulo_selecionado = st.sidebar.radio(
    "Escolha o Módulo:",
    options=[
        "📊 Produtividade de Oficinas", 
        "💶 Faturação e Orçamento Diário"
    ]
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

def calcular_dias_uteis_mes(ano, mes, cidade="Nacional (Sem Municipal)"):
    if mes == 12:
        proximo_mes = date(ano + 1, 1, 1)
    else:
        proximo_mes = date(ano, mes + 1, 1)
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

dias_uteis_calculados = calcular_dias_uteis_mes(ano_corrente, mes_selecionado, cidade_feriado)
decorridos_iniciais = dias_uteis_calculados if (mes_selecionado < mes_atual_num or ano_corrente < 2026) else min(8, dias_uteis_calculados)

dias_uteis_padrao_mes = st.sidebar.slider(
    f"Total de Dias Úteis ({meses_pt[mes_selecionado].capitalize()}):", 
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
                    match = df_perfis[df_perfis["Utilizador_clean"] == user_input]
                        
                    if not match.empty:
                        st.session_state.autenticado = True
                        st.session_state.utilizador = match.iloc[0]["Utilizador"]
                        permissoes_str = str(match.iloc[0]["Oficinas permitidas"]).strip()
                        
                        agrupar_val = str(match.iloc[0].get("Agrupar", "não")).strip().lower()
                        st.session_state.agrupar_perfil = (agrupar_val == "sim")
                        
                        agrupamento_str = str(match.iloc[0].get("Agrupamento", "")).strip()
                        if agrupamento_str and agrupamento_str.lower() != "nan":
                            st.session_state.oficinas_agrupadas_perfil = [o.strip().upper() for o in agrupamento_str.split(",") if o.strip()]
                        else:
                            st.session_state.oficinas_agrupadas_perfil = []

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
            st.session_state.agrupar_perfil = False
            st.session_state.oficinas_agrupadas_perfil = []
            st.rerun()

file_mecanicos_up = st.session_state.get("up_mec")
file_hf_up = st.session_state.get("up_hf")
file_ausencias_up = st.session_state.get("up_aus")
file_fact1_up = st.session_state.get("up_f1")
file_fact2_up = st.session_state.get("up_f2")
file_orc_up = st.session_state.get("up_orc")

path_mecanicos = file_mecanicos_up if file_mecanicos_up is not None else ("tabela mecânicos.xlsx" if os.path.exists("tabela mecânicos.xlsx") else None)

nome_arq_hf = f"Comparativo de Horas_{sigla_mes}_{ano_corrente}.xlsx"
path_hf = file_hf_up if file_hf_up is not None else (nome_arq_hf if os.path.exists(nome_arq_hf) else None)

nome_arq_aus = f"Listagem de AusenciasPresenças_{sigla_mes}_{ano_corrente}.xlsx"
path_ausencias = nome_arq_aus if os.path.exists(nome_arq_aus) else None

def converter_valor_pt(val_str):
    if pd.isna(val_str):
        return 0.0
    s = str(val_str).replace("€", "").strip()
    if not s:
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

def limpar_str(s):
    if pd.isna(s): return ""
    s_nfkd = unicodedata.normalize('NFKD', str(s))
    return "".join([c for c in s_nfkd if not unicodedata.combining(c)]).upper().strip()

def oficina_permitida_check(oficina_nome):
    if st.session_state.acesso_total:
        return True
    oficina_up = limpar_str(oficina_nome)
    for permitida in st.session_state.oficinas_autorizadas:
        p_clean = limpar_str(permitida)
        if p_clean == oficina_up or p_clean in oficina_up or oficina_up in p_clean:
            return True
    return False

def detetar_cidade(nome_oficina):
    u = limpar_str(nome_oficina)
    if " PORTALEGRE" in u or u.endswith("PORTALEGRE") or "PORTALEGRE " in u: return "Portalegre"
    if " BEJA" in u or u.endswith("BEJA") or "BEJA " in u: return "Beja"
    if " FARO" in u or u.endswith("FARO") or "FARO " in u: return "Faro"
    if " EVORA" in u or u.endswith("EVORA") or "EVORA " in u or "ÉVORA" in u: return "Évora"
    if " PORTIMAO" in u or u.endswith("PORTIMAO") or "PORTIMÃO" in u: return "Portimão"
    if " SETUBAL" in u or u.endswith("SETUBAL") or "SETÚBAL" in u: return "Setúbal"
    if " SINES" in u or u.endswith("SINES") or "SINES " in u: return "Sines"
    if " ALMADA" in u or u.endswith("ALMADA") or "ALMADA " in u: return "Almada"
    if " PALMELA" in u or u.endswith("PALMELA") or "PALMELA " in u: return "Palmela"
    if " MONTIJO" in u or u.endswith("MONTIJO") or "MONTIJO " in u: return "Montijo"
    return "Geral"

def padronizar_oficina_fat(nome):
    u = limpar_str(nome)
    
    if "PORTALEGRE" in u:
        if "EUROREPAR" in u:
            return "Eurorepar Portalegre"
        elif "PEUGEOT" in u:
            return "Peugeot Portalegre"
            
    if "EVORA" in u or "ÉVORA" in u:
        if "COLISAO" in u or "COLISÃO" in u:
            return "Colisão Évora"
        elif "EUROREPAR" in u:
            return "Eurorepar Évora"
        elif "RENAULT" in u:
            return "Mecânica Renault Évora"
        else:
            return "Mecânica Stellantis Évora"

    cidades_lista = ["BEJA", "FARO", "PORTIMAO", "SETUBAL", "SINES", "ALMADA", "PALMELA", "MONTIJO"]
    cidade_encontrada = next((c for c in cidades_lista if c in u), None)
    
    if cidade_encontrada:
        if "COLISAO" in u or "COLISÃO" in u:
            return f"Colisão {cidade_encontrada.title()}"
        else:
            return f"Mecânica {cidade_encontrada.title()}"
            
    # Limpa prefixos residuais do tipo "OF. " para evitar duplicados
    limpo = str(nome).strip()
    if limpo.upper().startswith("OF. "):
        limpo = limpo[4:].strip()
    elif limpo.upper().startswith("OF "):
        limpo = limpo[3:].strip()
    return limpo

# ==========================================
# CORPO PRINCIPAL
# ==========================================
if not st.session_state.autenticado:
    st.warning("🔒 Por favor, faça login na barra lateral para aceder aos relatórios.")
elif not path_mecanicos or not path_hf:
    with container_filtros:
        st.header("Filtros")
        st.caption("Aguardando ficheiros de dados.")
    st.warning(f"⚠️ Não foram encontrados os ficheiros padrão no repositório para o período de {meses_pt[mes_selecionado].capitalize()} {ano_corrente} ({sigla_mes}). Certifique-se de que os carregou para o GitHub ou use o painel de admin.")
else:
    if modulo_selecionado == "📊 Produtividade de Oficinas":
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

        def normalizar_oficina(nome_oficina):
            if pd.isna(nome_oficina):
                return "Geral"
            nome_str = str(nome_oficina).strip()
            nome_upper = nome_str.upper()
            
            for oficina_oficial in oficinas_oficiais:
                if oficina_oficial.upper() == nome_upper:
                    return oficina_oficial
                    
            if "BEJA" in nome_upper: return "MECÂNICA BEJA"
            if "PORTIMÃO" in nome_upper or "PORTIMAO" in nome_upper: return "MECÂNICA PORTIMÃO"
            if "FARO" in nome_upper: return "MECÂNICA FARO"
            if "ÉVORA" in nome_upper or "EVORA" in nome_upper: return "MECÂNICA ÉVORA"
            if "SETÚBAL" in nome_upper or "SETUBAL" in nome_upper: return "MECÂNICA SETÚBAL"
            if "SINES" in nome_upper: return "MECÂNICA SINES"
            if "PORTALEGRE" in nome_upper: return "PORTALEGRE"
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
            "Técnico", "Local", "Oficina", "Função", 
            "Horas potencial dia", "Horas reais dia", "Ausencias",
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
                        if "BEJA" in of_upper: local_val = "Beja"
                        elif "PORTIMÃO" in of_upper or "PORTIMAO" in of_upper: local_val = "Portimão"
                        elif "ÉVORA" in of_upper or "EVORA" in of_upper: local_val = "Évora"
                        elif "FARO" in of_upper: local_val = "Faro"
                        elif "SETÚBAL" in of_upper or "SETUBAL" in of_upper: local_val = "Setúbal"
                        elif "SINES" in of_upper: local_val = "Sines"
                        elif "PORTALEGRE" in of_upper: local_val = "Portalegre"

                    linhas_sem_produtivo.append({
                        "Técnico": "Sem produtivo",
                        "Local": local_val,
                        "Oficina": oficina_val,
                        "Função": "Sem Produtivo",
                        "Horas potencial dia": 0.0,
                        "Horas reais dia": 0.0,
                        "Ausencias": 0.0,
                        "Horas disponiveis": 0.0,
                        "Horas facturadas": horas_fact_val,
                        "Produtividade %": 0.0
                    })

        df_sem_prod_final = pd.DataFrame(linhas_sem_produtivo)
        df_completo = pd.concat([df_exibicao_tecnicos, df_sem_prod_final], ignore_index=True) if not df_sem_prod_final.empty else df_exibicao_tecnicos

        df_completo = df_completo[df_completo["Oficina"].apply(oficina_permitida_check)]

        with container_filtros:
            st.header("Filtros")
            locais_unicos = sorted(df_completo["Local"].dropna().astype(str).unique().tolist())
            locais_disponiveis = (["Todas"] + locais_unicos) if len(locais_unicos) > 1 else locais_unicos
            sel_local = st.selectbox("Selecionar Cidade (Local)", locais_disponiveis if locais_disponiveis else ["Sem acesso"])

            df_filtrado_local = df_completo[df_completo["Local"] == sel_local] if sel_local != "Todas" else df_completo

            oficinas_unicas = sorted(df_filtrado_local["Oficina"].dropna().astype(str).unique().tolist())
            oficinas_disponiveis = (["Todas"] + oficinas_unicas) if len(oficinas_unicas) > 1 else oficinas_unicas
            sel_oficina = st.selectbox("Selecionar Oficina", oficinas_disponiveis if oficinas_disponiveis else ["Sem acesso"])

        df_final_filtro = df_filtrado_local[df_filtrado_local["Oficina"] == sel_oficina] if sel_oficina != "Todas" else df_filtrado_local

        tot_h_potencial = df_final_filtro["Horas potencial dia"].sum() if "Horas potencial dia" in df_final_filtro.columns else 0.0
        tot_h_reais = df_final_filtro["Horas reais dia"].sum() if "Horas reais dia" in df_final_filtro.columns else 0.0
        tot_ausencias = df_final_filtro["Ausencias"].sum() if "Ausencias" in df_final_filtro.columns else 0.0
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
            "Ausencias": tot_ausencias,
            "Horas disponiveis": tot_h_disp,
            "Horas facturadas": tot_h_fact,
            "Produtividade %": tot_prod
        }

        df_exibicao_final = pd.concat([df_final_filtro, pd.DataFrame([linha_total])], ignore_index=True)

        def destacar_total(s):
            is_total = s.index.isin(df_exibicao_final[df_exibicao_final["Técnico"] == "TOTAL"].index) if "Técnico" in df_exibicao_final.columns else []
            return ['font-weight: bold; background-color: #f0f2f6;' if v else '' for v in is_total]

        st.subheader(f"Mapa de Produtividade - {meses_pt[mes_selecionado].capitalize()} {ano_corrente}")
        st.dataframe(
            df_exibicao_final.style.apply(destacar_total, axis=1).format({
                "Horas potencial dia": "{:.1f}",
                "Horas reais dia": "{:.1f}",
                "Ausencias": "{:.1f}",
                "Horas disponiveis": "{:.1f}",
                "Horas facturadas": "{:.2f}",
                "Produtividade %": "{:.1f}%"
            }, na_rep="-"),
            use_container_width=True
        )

    elif modulo_selecionado == "💶 Faturação e Orçamento Diário":
        st.subheader(f"Faturação e Orçamento Diário - {meses_pt[mes_selecionado].capitalize()} {ano_corrente}")
        
        excel_orc_path = "Orçamento-agrupamento oficinas.xlsx"
        df_orc = pd.DataFrame()
        
        target_orc_file = file_orc_up if file_orc_up is not None else (excel_orc_path if os.path.exists(excel_orc_path) else None)
        
        if target_orc_file is not None:
            try:
                df_orc = pd.read_excel(target_orc_file, sheet_name="Oficina_ORÇAMENTO")
            except Exception as e:
                st.error(f"Erro ao ler o ficheiro de orçamentos: {e}")
        
        df_fact_total = pd.DataFrame()
        csv_fontes = [file_fact1_up, file_fact2_up]
        
        if all(f is None for f in csv_fontes):
            arquivos_csv_pasta = glob.glob("*.csv")
            for arq in arquivos_csv_pasta:
                if "MULTIAUTO" in arq.upper() or "FACTURACAO" in arq.upper() or "MAPA" in arq.upper():
                    csv_fontes.append(arq)
            if not csv_fontes or all(f is None for f in csv_fontes):
                csv_fontes = arquivos_csv_pasta

        for file_f in csv_fontes:
            if file_f is not None:
                try:
                    if hasattr(file_f, "getvalue"):
                        file_bytes = io.BytesIO(file_f.getvalue())
                    elif hasattr(file_f, "read"):
                        file_bytes = file_f.read()
                    else:
                        file_bytes = file_f

                    if hasattr(file_bytes, "seek"):
                        file_bytes.seek(0)
                        
                    conteudo = file_bytes.read().decode("iso-8859-1", errors="ignore") if hasattr(file_bytes, "read") else open(file_bytes, encoding="iso-8859-1", errors="ignore").read()
                    linhas = conteudo.splitlines()
                    
                    dados_temp = []
                    for linha in linhas:
                        if not linha.strip(): continue
                        partes = []
                        for sep in [";", "\t", ","]:
                            p = [x.strip().replace('"', '') for x in linha.split(sep)]
                            if len(p) > 5:
                                partes = p
                                break
                        if len(partes) > 0:
                            dados_temp.append(partes)
                            
                    if dados_temp:
                        t_temp = pd.DataFrame(dados_temp)
                        df_fact_total = pd.concat([df_fact_total, t_temp], ignore_index=True)
                except Exception as e:
                    pass
                
        if df_fact_total.empty:
            st.warning("⚠️ Não foram encontrados ficheiros CSV de faturação na pasta do repositório.")
        else:
            idx_ofic = 0   # Coluna A
            idx_val = 24   # Coluna Y (Total Liqu)
            idx_data = 28  # Coluna AC (Data Factura)

            df_fact_total["Oficina_CSV"] = df_fact_total.iloc[:, idx_ofic].astype(str).str.strip()
            
            invalidos = ["NAN", "", "TOTAL", "MAPA DE FACTURAÇÃO"]
            df_fact_total = df_fact_total[~df_fact_total["Oficina_CSV"].str.upper().isin(invalidos)]
            df_fact_total = df_fact_total[~df_fact_total["Oficina_CSV"].str.startswith("*")]

            col_data_raw = df_fact_total.iloc[:, idx_data]
            df_fact_total["Data_Parsed"] = pd.to_datetime(col_data_raw, dayfirst=True, errors="coerce")

            col_valor_raw = df_fact_total.iloc[:, idx_val]
            df_fact_total["Valor_Num"] = col_valor_raw.apply(converter_valor_pt)

            df_fact_total["Oficina_Agrupada"] = df_fact_total["Oficina_CSV"].apply(padronizar_oficina_fat)

            oficinas_disponiveis_fat = sorted(list(set([o for o in df_fact_total["Oficina_Agrupada"].dropna().astype(str).unique() if o.strip() and "MAPA" not in o.upper() and not set(o.strip()).issubset({"*"})])))
            if not df_orc.empty:
                df_orc.columns = [str(c).strip() for c in df_orc.columns]
                oficinas_orc_unicas = sorted([o for o in df_orc[df_orc.columns[0]].dropna().astype(str).unique() if o.strip() and "MAPA" not in o.upper() and not set(o.strip()).issubset({"*"})])
                oficinas_disponiveis_fat = sorted(list(set(oficinas_disponiveis_fat + oficinas_orc_unicas)))
            
            if st.session_state.acesso_total:
                oficinas_autorizadas_lista = oficinas_disponiveis_fat
            else:
                if st.session_state.agrupar_perfil and st.session_state.oficinas_agrupadas_perfil:
                    oficinas_autorizadas_lista = st.session_state.oficinas_agrupadas_perfil
                else:
                    oficinas_autorizadas_lista = [o for o in oficinas_disponiveis_fat if oficina_permitida_check(o)]
            
            if not oficinas_autorizadas_lista:
                oficinas_autorizadas_lista = ["Sem acesso"]

            col_sel_cidade, col_sel_ofic = st.columns(2)
            
            cidades_autorizadas = sorted(list(set([detetar_cidade(o) for o in oficinas_autorizadas_lista])))

            with col_sel_cidade:
                cidade_selecionada = st.selectbox("Selecionar Cidade:", options=cidades_autorizadas)

            if cidade_selecionada != "Geral":
                if not st.session_state.acesso_total and st.session_state.agrupar_perfil and st.session_state.oficinas_agrupadas_perfil:
                    oficinas_filtradas_cidade = [o for o in st.session_state.oficinas_agrupadas_perfil if detetar_cidade(o) == cidade_selecionada or cidade_selecionada.upper() in o.upper()]
                else:
                    oficinas_filtradas_cidade = [o for o in oficinas_autorizadas_lista if detetar_cidade(o) == cidade_selecionada]
                
                if not oficinas_filtradas_cidade:
                    oficinas_filtradas_cidade = [o for o in oficinas_autorizadas_lista if cidade_selecionada.upper() in o.upper()]
            else:
                oficinas_filtradas_cidade = oficinas_autorizadas_lista

            with col_sel_ofic:
                oficina_selecionada_fat = st.selectbox("Selecionar Oficina (Faturação):", options=oficinas_filtradas_cidade)
                
            sel_clean = limpar_str(oficina_selecionada_fat)
            
            df_oficina_fat = df_fact_total[
                (df_fact_total["Oficina_Agrupada"].apply(limpar_str) == sel_clean) |
                (df_fact_total["Oficina_CSV"].apply(limpar_str) == sel_clean) |
                df_fact_total["Oficina_Agrupada"].apply(lambda x: sel_clean in limpar_str(x))
            ]
            
            df_mes_ano = df_oficina_fat[(df_oficina_fat["Data_Parsed"].dt.month == mes_selecionado) & (df_oficina_fat["Data_Parsed"].dt.year == ano_corrente)]
            if not df_mes_ano.empty:
                df_oficina_fat = df_mes_ano
            
            df_diario = df_oficina_fat.groupby(df_oficina_fat["Data_Parsed"].dt.date, as_index=False)["Valor_Num"].sum()
            df_diario.rename(columns={"Valor_Num": "Facturacao_Diaria", "Data_Parsed": "Data"}, inplace=True)
            df_diario = df_diario.dropna(subset=["Data"]).sort_values("Data")
            
            inicio_mes = date(ano_corrente, mes_selecionado, 1)
            if mes_selecionado == 12:
                fim_mes = date(ano_corrente + 1, 1, 1) - timedelta(days=1)
            else:
                fim_mes = date(ano_corrente, mes_selecionado + 1, 1) - timedelta(days=1)
                
            feriados_mes = obter_feriados_mes(ano_corrente, mes_selecionado, cidade_feriado)
            
            dias_uteis_lista = []
            atual = inicio_mes
            while atual <= fim_mes:
                if atual.weekday() < 5 and atual not in feriados_mes:
                    dias_uteis_lista.append(atual)
                atual += timedelta(days=1)
                
            df_completo_dias = pd.DataFrame({"Data": dias_uteis_lista})
            df_completo_dias = pd.merge(df_completo_dias, df_diario, on="Data", how="left").fillna(0.0)
            
            df_completo_dias["Data_Str"] = pd.to_datetime(df_completo_dias["Data"]).dt.strftime("%d/%m/%Y")
            
            hoje_atual = datetime.now().date()
            if ano_corrente == hoje_atual.year and mes_selecionado == hoje_atual.month:
                data_limite_grafico = hoje_atual
            else:
                data_limite_grafico = fim_mes

            df_completo_dias["Facturacao_Acumulada"] = df_completo_dias["Facturacao_Diaria"].cumsum()
            
            df_completo_dias["Facturacao_Acumulada_Graf"] = df_completo_dias.apply(
                lambda row: row["Facturacao_Acumulada"] if row["Data"] <= data_limite_grafico else None, axis=1
            )
            df_completo_dias["Facturacao_Diaria_Graf"] = df_completo_dias.apply(
                lambda row: row["Facturacao_Diaria"] if row["Data"] <= data_limite_grafico else 0.0, axis=1
            )

            orcamento_mensal = 0.0
            if not df_orc.empty and len(df_orc.columns) >= 5:
                col_ofic = df_orc.columns[0]
                col_mes = df_orc.columns[3]
                col_val = df_orc.columns[4]
                
                nome_mes_alvo = limpar_str(meses_pt[mes_selecionado])
                exclusoes = ["TOTAL", "GRUPO", "MAPA", "ACUMULADO", "GERAL", "RESUMO", "TODAS"]
                
                orc_encontrado = 0.0
                for _, row in df_orc.iterrows():
                    ofic_row = limpar_str(row[col_ofic])
                    
                    if not ofic_row or any(exc in ofic_row for exc in exclusoes):
                        continue
                        
                    val_mes_cell = row[col_mes]
                    mes_encontrado = False
                    if isinstance(val_mes_cell, (pd.Timestamp, datetime, date)):
                        if val_mes_cell.month == mes_selecionado:
                            mes_encontrado = True
                    else:
                        mes_row_str = limpar_str(val_mes_cell)
                        if nome_mes_alvo in mes_row_str or mes_row_str in nome_mes_alvo:
                            mes_encontrado = True
                        else:
                            try:
                                if int(val_mes_cell) == mes_selecionado:
                                    mes_encontrado = True
                            except:
                                pass
                                
                    if mes_encontrado:
                        if st.session_state.acesso_total:
                            if sel_clean == ofic_row:
                                orc_encontrado = converter_valor_pt(row[col_val])
                                break
                        else:
                            if sel_clean == ofic_row or sel_clean in ofic_row or ofic_row in sel_clean:
                                val_v = converter_valor_pt(row[col_val])
                                if val_v > 0:
                                    orc_encontrado = val_v
                                    break
                                    
                orcamento_mensal = orc_encontrado
                                
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
                st.metric(label="ORÇAMENTO MENSAL", value=f"{orcamento_mensal:,.0f} €".replace(",", " "))
                st.metric(label="ORÇAMENTO DIÁRIO", value=f"{orcamento_diario:,.0f} €".replace(",", " "))
                st.metric(label="FATURAÇÃO À DATA", value=f"{facturacao_a_data:,.0f} €".replace(",", " "))
                
                cor_desvio = "green" if desvio_orcamento >= 0 else "red"
                desvio_formatado = f"{desvio_orcamento:,.0f} €".replace(",", " ")
                
                st.markdown(f"""
                    <div style="font-size: 14px; font-weight: 600; color: rgb(49, 51, 63); margin-bottom: 0px;">DESVIO ORÇAMENTO</div>
                    <div style="font-size: 24px; font-weight: 700; color: {cor_desvio}; margin-top: 0px; margin-bottom: 0px;">{desvio_formatado}</div>
                """, unsafe_allow_html=True)
                
                st.metric(label="", value="", delta=f"{desvio_orcamento:,.0f} €".replace(",", " "), label_visibility="collapsed")
