import pandas as pd
import streamlit as st
from datetime import datetime, date, timedelta
import os
import glob
import altair as alt

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
    format_func=lambda x: meses_pt[x], 
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
        date(ano, 1, 1),   
        date(ano, 4, 25),  
        date(ano, 5, 1),   
        date(ano, 6, 10),  
        date(ano, 8, 15),  
        date(ano, 10, 5),  
        date(ano, 12, 1),  
        date(ano, 12, 8),  
        date(ano, 12, 25), 
    ]
    
    if ano == 2026:
        feriados.extend([
            date(2026, 4, 3),   
            date(2026, 6, 4),   
        ])
        
    cidade_upper = str(cidade).upper()
    if "BEJA" in cidade_upper:
        feriados.append(date(ano, 5, 14))    
    elif "ÉVORA" in cidade_upper or "EVORA" in cidade_upper:
        feriados.append(date(ano, 6, 29))    
    elif "PORTIMÃO" in cidade_upper or "PORTIMAO" in cidade_upper:
        feriados.append(date(ano, 12, 19))   
    elif "FARO" in cidade_upper:
        feriados.append(date(ano, 9, 7))     
    elif "SETÚBAL" in cidade_upper or "SETUBAL" in cidade_upper:
        feriados.append(date(ano, 9, 15))    
    elif "SINES" in cidade_upper:
        feriados.append(date(ano, 11, 24))   
        
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
file_fact1_up = None
file_fact2_up = None
file_orc_up = None

if tem_permissao_upload:
    st.sidebar.markdown("---")
    st.sidebar.header("Atualizar Ficheiros (Admin)")
    file_mecanicos_up = st.sidebar.file_uploader("Substituir 'tabela mecânicos'", type=["xlsx"])
    file_hf_up = st.sidebar.file_uploader("Substituir 'Comparativo de Horas'", type=["xlsx"])
    file_ausencias_up = st.sidebar.file_uploader("Substituir 'Listagem de Ausencias'", type=["xlsx"])
    file_fact1_up = st.sidebar.file_uploader("Mapa de Faturação 1 (CSV)", type=["csv"])
    file_fact2_up = st.sidebar.file_uploader("Mapa de Faturação 2 (CSV)", type=["csv"])
    file_orc_up = st.sidebar.file_uploader("Orçamento e Agrupamentos (XLSX)", type=["xlsx"])

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

if file_mecanicos_up is not None:
    path_mecanicos = file_mecanicos_up
elif os.path.exists("tabela mecânicos.xlsx"):
    path_mecanicos = "tabela mecânicos.xlsx"
else:
    path_mecanicos = None

nome_arq_hf = f"Comparativo de Horas_{sigla_mes}_{ano_corrente}.xlsx"
if file_hf_up is not None:
    path_hf = file_hf_up
elif os.path.exists(nome_arq_hf):
    path_hf = nome_arq_hf
else:
    path_hf = None

nome_arq_aus = f"Listagem de AusenciasPresenças_{sigla_mes}_{ano_corrente}.xlsx"
path_ausencias = nome_arq_aus if os.path.exists(nome_arq_aus) else None

# ==========================================
# CORPO PRINCIPAL
# ==========================================
if not st.session_state.autenticado:
    st.warning("🔒 Por favor, faça login na barra lateral para aceder aos relatórios.")
elif not path_mecanicos or not path_hf:
    with container_filtros:
        st.header("Filtros")
        st.caption("Aguardando ficheiros de dados.")
    st.warning(f"⚠️ Não foram encontrados os ficheiros padrão no repositório para o período de {meses_pt[mes_selecionado]} {ano_corrente} ({sigla_mes}). Certifique-se de que os carregou para o GitHub ou use o painel de admin.")
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

    elif modulo_selecionado == "💶 Faturação e Orçamento Diário":
        st.subheader(f"Faturação e Orçamento Diário - {meses_pt[mes_selecionado]} {ano_corrente}")
        
        excel_orc_path = "Orçamento-agrupamento oficinas.xlsx"
        df_orc = pd.DataFrame()
        df_agrup = pd.DataFrame()
        
        if file_orc_up is not None:
            try:
                df_orc = pd.read_excel(file_orc_up, sheet_name="Oficina_ORÇAMENTO")
                df_agrup = pd.read_excel(file_orc_up, sheet_name="Agrupamento_oficinas")
            except Exception as e:
                st.error(f"Erro ao ler o ficheiro carregado de orçamentos: {e}")
        elif os.path.exists(excel_orc_path):
            try:
                df_orc = pd.read_excel(excel_orc_path, sheet_name="Oficina_ORÇAMENTO")
                df_agrup = pd.read_excel(excel_orc_path, sheet_name="Agrupamento_oficinas")
            except Exception as e:
                st.error(f"Erro ao ler o ficheiro de orçamentos e agrupamentos: {e}")
        
        df_fact_total = pd.DataFrame()
        
        # Leitura flexível com skiprows ajustável ou deteção automática
        for file_f in [file_fact1_up, file_fact2_up]:
            if file_f is not None:
                try:
                    file_f.seek(0)
                    t_temp = pd.read_csv(file_f, sep=";", encoding="iso-8859-1", skiprows=5)
                    if len(t_temp.columns) < 3:
                        file_f.seek(0)
                        t_temp = pd.read_csv(file_f, sep=",", encoding="iso-8859-1", skiprows=5)
                    if not t_temp.empty:
                        df_fact_total = pd.concat([df_fact_total, t_temp], ignore_index=True)
                except Exception as e:
                    pass

        if df_fact_total.empty:
            csv_fact_files = glob.glob(f"*FACTURAÇÃO*{ano_corrente}_{mes_selecionado:02d}.csv")
            if not csv_fact_files:
                csv_fact_files = glob.glob("*FACTURAÇÃO*.csv")
                
            for f in csv_fact_files:
                try:
                    temp_df = pd.read_csv(f, sep=";", encoding="iso-8859-1", skiprows=5)
                    if not temp_df.empty:
                        df_fact_total = pd.concat([df_fact_total, temp_df], ignore_index=True)
                except Exception as e:
                    pass
                
        if df_fact_total.empty:
            st.warning("⚠️ Não foram encontrados mapas de faturação. Carregue os ficheiros CSV de faturação na barra lateral (Admin).")
        else:
            # Deteção ultra-robusta de colunas de oficina, data e valor
            col_oficina_f = df_fact_total.columns[0]
            
            col_data_f = next((c for c in df_fact_total.columns if any(k in str(c).lower() for k in ["data", "factura", "emissão", "dia"])), df_fact_total.columns[min(28, len(df_fact_total.columns)-1)])
            col_valor_f = next((c for c in df_fact_total.columns if any(k in str(c).lower() for k in ["liquido", "líquido", "valor", "total", "s/iva", "euros"])), df_fact_total.columns[min(24, len(df_fact_total.columns)-1)])
            
            df_fact_total["Oficina_CSV"] = df_fact_total[col_oficina_f].astype(str).str.strip()
            
            # Conversão robusta de datas (com dayfirst=True para formato europeu DD-MM-YYYY)
            df_fact_total["Data_Parsed"] = pd.to_datetime(df_fact_total[col_data_f], dayfirst=True, errors="coerce")

            # Se ainda houver muitas datas NaT, tenta parse universal misto
            if df_fact_total["Data_Parsed"].isna().mean() > 0.5:
                df_fact_total["Data_Parsed"] = pd.to_datetime(df_fact_total[col_data_f], format="mixed", dayfirst=True, errors="coerce")

            # Tratamento numérico robusto
            val_series = df_fact_total[col_valor_f]
            if val_series.dtype == object:
                df_fact_total["Valor_Num"] = val_series.astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
                df_fact_total["Valor_Num"] = pd.to_numeric(df_fact_total["Valor_Num"], errors="coerce").fillna(0.0)
            else:
                df_fact_total["Valor_Num"] = pd.to_numeric(val_series, errors="coerce").fillna(0.0)

            if not df_agrup.empty:
                col_agrup_oficina = df_agrup.columns[0]
                col_agrup_destino = df_agrup.columns[3] if len(df_agrup.columns) > 3 else df_agrup.columns[1]
                
                dict_agrupamento = dict(zip(df_agrup[col_agrup_oficina].astype(str).str.strip().str.upper(), df_agrup[col_agrup_destino].astype(str).str.strip()))
                df_fact_total["Oficina_Agrupada"] = df_fact_total["Oficina_CSV"].str.upper().map(dict_agrupamento).fillna(df_fact_total["Oficina_CSV"])
            else:
                df_fact_total["Oficina_Agrupada"] = df_fact_total["Oficina_CSV"]

            oficinas_disponiveis_fat = sorted(df_fact_total["Oficina_Agrupada"].unique().tolist())
            if not df_orc.empty:
                oficinas_orc_unicas = sorted(df_orc["Oficina"].astype(str).str.strip().unique().tolist())
                oficinas_disponiveis_fat = sorted(list(set(oficinas_disponiveis_fat + oficinas_orc_unicas)))
            
            col_sel_ofic, col_empty = st.columns([2, 2])
            with col_sel_ofic:
                oficina_selecionada_fat = st.selectbox("Selecionar Oficina (Faturação):", options=oficinas_disponiveis_fat)
                
            df_oficina_fat = df_fact_total[df_fact_total["Oficina_Agrupada"].astype(str).str.upper() == oficina_selecionada_fat.upper()]
            
            df_diario = df_oficina_fat.groupby(df_oficina_fat["Data_Parsed"].dt.date, as_index=False)["Valor_Num"].sum()
            df_diario.rename(columns={"Valor_Num": "Facturacao_Diaria", "Data_Parsed": "Data"}, inplace=True)
            df_diario = df_diario.dropna(subset=["Data"]).sort_values("Data")
            
            inicio_mes = date(ano_corrente, mes_selecionado, 1)
            if mes_selecionado == 12:
                fim_mes = date(ano_corrente + 1, 1, 1) - timedelta(days=1)
            else:
                fim_mes = date(ano_corrente, mes_selecionado + 1, 1) - timedelta(days=1)
                
            range_dias = pd.date_range(start=inicio_mes, end=fim_mes)
            df_completo_dias = pd.DataFrame({"Data": range_dias.date})
            df_completo_dias = pd.merge(df_completo_dias, df_diario, on="Data", how="left").fillna(0.0)
            df_completo_dias["Facturacao_Acumulada"] = df_completo_dias["Facturacao_Diaria"].cumsum()
            
            orcamento_mensal = 0.0
            if not df_orc.empty:
                df_orc["Data_Parsed"] = pd.to_datetime(df_orc["Data"], format="mixed", errors="coerce")
                match_orc = df_orc[
                    (df_orc["Oficina"].astype(str).str.strip().str.upper() == oficina_selecionada_fat.upper()) & 
                    (df_orc["Data_Parsed"].dt.year == ano_corrente) & 
                    (df_orc["Data_Parsed"].dt.month == mes_selecionado)
                ]
                if not match_orc.empty:
                    orcamento_mensal = float(match_orc.iloc[0]["Valor"])
                    
            orcamento_diario = orcamento_mensal / dias_uteis_padrao_mes if dias_uteis_padrao_mes > 0 else 0.0
            
            orcamentos_ac = []
            acum = 0.0
            for d in df_completo_dias["Data"]:
                if d.weekday() < 5:
                    acum += orcamento_diario
                orcamentos_ac.append(acum)
            df_completo_dias["Orcamento_Acumulado"] = orcamentos_ac
            
            facturacao_a_data = df_completo_dias["Facturacao_Acumulada"].iloc[-1]
            desvio_orcamento = facturacao_a_data - orcamento_mensal
            
            col_graf, col_kpi = st.columns([3, 1])
            
            with col_graf:
                base = alt.Chart(df_completo_dias).encode(
                    x=alt.X('Data:T', title='Data', axis=alt.Axis(format='%d/%m/%Y', labelAngle=-45))
                )
                
                barras = base.mark_bar(color='#115e59').encode(
                    y=alt.Y('Facturacao_Acumulada:Q', title='Valor (€)'),
                    tooltip=['Data', 'Facturacao_Diaria', 'Facturacao_Acumulada']
                )
                
                linha_orc = base.mark_line(color='#f97316', strokeWidth=3).encode(
                    y='Orcamento_Acumulado:Q',
                    tooltip=['Data', 'Orcamento_Acumulado']
                )
                
                chart = (barras + linha_orc).properties(height=400).interactive()
                st.altair_chart(chart, use_container_width=True)
                
            with col_kpi:
                st.metric(label="ORÇAMENTO MENSAL", value=f"{orcamento_mensal:,.2f} €".replace(",", " ").replace(".", ","))
                st.metric(label="ORÇAMENTO DIÁRIO", value=f"{orcamento_diario:,.2f} €".replace(",", " ").replace(".", ","))
                st.metric(label="FATURAÇÃO À DATA", value=f"{facturacao_a_data:,.2f} €".replace(",", " ").replace(".", ","))
                st.metric(label="DESVIO ORÇAMENTO", value=f"{desvio_orcamento:,.2f} €".replace(",", " ").replace(".", ","), delta=f"{desvio_orcamento:,.2f} €")
