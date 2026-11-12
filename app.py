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

# 3. SECÇÃO: CARREGAR FICHEIROS EXCEL (APENas PARA VAZUL E MCHEGADINHO)
utilizadores_com_permissao_upload = ["vazul", "mchegadinho"]
utilizador_atual = st.session_state.utilizador.strip().lower()
tem_permissao_upload = st.session_state.autenticado and (utilizador_atual in [u.lower() for u in utilizadores_com_permissao_upload])

file_perfis = None
file_mecanicos = None
file_hf = None
file_ausencias = None

if tem_permissao_upload:
    st.sidebar.markdown("---")
    st.sidebar.header("Carregar Ficheiros Excel")
    file_perfis = st.sidebar.file_uploader("Carregar 'perfis_utilizadores.xlsx'", type=["xlsx"])
    file_mecanicos = st.sidebar.file_uploader("Carregar 'tabela mecânicos.xlsx'", type=["xlsx"])
    file_hf = st.sidebar.file_uploader("Carregar 'Comparativo de Horas'", type=["xlsx"])
    file_ausencias = st.sidebar.file_uploader("Carregar 'Listagem de Ausencias'", type=["xlsx"])

# LÓGICA DE AUTENTICAÇÃO E LOGOUT
with container_login:
    st.header("Acesso / Login")
    # Para validar o login, precisamos de ler temporariamente o ficheiro de perfis se o utilizador for admin, 
    # mas como os admins precisam de carregar os ficheiros, vamos permitir carregar o perfis_utilizadores mesmo antes do login completo, 
    # ou assumir que o perfis_utilizadores está acessível. Vamos manter o uploader de perfis visível para todos no login se necessário, 
    # ou restringir os outros. 
    # Nota: para fazer login, precisamos do df_perfis. Vamos garantir que o uploader de perfis aparece no login:
    
if not st.session_state.autenticado:
    with container_login:
        file_perfis_temp = st.file_uploader("Carregar 'perfis_utilizadores.xlsx' (Login)", type=["xlsx"], key="login_perfis")
        if file_perfis_temp:
            file_perfis = file_perfis_temp
            df_perfis = pd.read_excel(file_perfis)
            df_perfis["Utilizador_clean"] = df_perfis["Utilizador"].astype(str).str.strip().str.lower()
            df_perfis["password_clean"] = df_perfis["password"].astype(str).str.strip()

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
            st.caption("Carrege o ficheiro 'perfis_utilizadores.xlsx' para iniciar sessão.")
else:
    with container_login:
        st.success(f"Sessão iniciada como: **{st.session_state.utilizador}**")
        if st.button("Logout / Mudar Utilizador"):
            st.session_state.autenticado = False
            st.session_state.utilizador = ""
            st.session_state.oficinas_autorizadas = []
            st.session_state.acesso_total = False
            st.rerun()

# PROCESSAMENTO DOS DADOS (Se for admin, usa os ficheiros carregados; se for outro utilizador, assume que os ficheiros base estão disponíveis ou tratados)
# Nota: Para os utilizadores normais gerarem o relatório, eles precisam que os ficheiros de mecânicos, horas e ausências estejam disponíveis. 
# Se os admins (Vazul / mchegadinho) carregarem os ficheiros, podes guardá-los ou mantê-los na sessão, ou exigir que os admins façam o upload.

if not st.session_state.autenticado:
    st.warning("🔒 Por favor, faça login no topo da barra lateral para aceder aos relatórios.")
elif not tem_permissao_upload and not file_mecanicos:
    st.info("ℹ️ A aguardar que o administrador carregue os ficheiros de dados da oficina.")
elif file_mecanicos and file_hf:
    xl = pd.ExcelFile(file_mecanicos)
    df_mecanicos = xl.parse("Mecânicos" if "Mecânicos" in xl.sheet_names else xl.sheet_names[0])
    df_hf = pd.read_excel(file_hf)
    
    # PROCESSAMENTO DE AUSÊNCIAS (TUDO O QUE NÃO É "Presença" NA COLUNA B)
    df_ausencias = pd.read_excel(file_ausencias) if file_ausencias else pd.DataFrame()
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

    col_operario_hf = next((c for c in ["Operário", "Operario", "Técnico", "Tecnico", "Nome", "Mecânico"] if c in df_hf.columns), df_hf.select_dtypes(include=['object']).columns[0])
    col_fact = next((c for c in ["Tempos facturados", "Tempos facturados ", "Horas facturadas", "Facturado", "Valor"] if c in df_hf.columns), df_hf.select_dtypes(include=['number']).columns[-1])
    col_oficina_hf = next((c for c in ["Oficina", "Escritório", "Escritorio", "Local"] if c in df_hf.columns), "Oficina")

    df_hf["Operario_clean"] = df_hf[col_operario_hf].astype(str).str.strip().str.upper()
    df_hf["Oficina_Normalizada"] = df_hf[col_oficina_hf].apply(normalizar_oficina)

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

    if not st.session_state.acesso_total:
        def oficina_permitida(oficina):
            oficina_up = str(oficina).upper()
            for permitida in st.session_state.oficinas_autorizadas:
                if permitida in oficina_up or oficina_up in permitida:
                    return True
            return False
        df_completo = df_completo[df_completo["Oficina"].apply(oficina_permitida)]

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

    st.subheader(f"Mapa de Produtividade - {meses_pt[mes_selecionado]} {ano_corrente}")
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
elif st.session_state.autenticado and tem_permissao_upload:
    with container_filtros:
        st.header("Filtros")
        st.caption("Carregue os ficheiros em cima para ativar os filtros.")
    st.info("👈 Por favor, carregue a **tabela de mecânicos** e o **comparativo de horas** na barra lateral para gerar o relatório.")
