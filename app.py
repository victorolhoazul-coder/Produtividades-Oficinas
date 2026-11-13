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
        
        # 1. Carregamento de Orçamentos e Agrupamentos
        excel_orc_path = "Orçamento-agrupamento oficinas.xlsx"
        df_orc, df_agrup = pd.DataFrame(), pd.DataFrame()
        target_orc = file_orc_up if file_orc_up is not None else (excel_orc_path if os.path.exists(excel_orc_path) else None)
        
        if target_orc is not None:
            try:
                df_orc = pd.read_excel(target_orc, sheet_name="Oficina_ORÇAMENTO")
                df_agrup = pd.read_excel(target_orc, sheet_name="Agrupamento_oficinas")
            except Exception as e:
                st.error(f"Erro ao ler orçamentos: {e}")
        
        # 2. Leitura dos CSVs de Faturação (saltando as 3 primeiras linhas de cabeçalho inútil)
        df_fact_total = pd.DataFrame()
        csv_fontes = [file_fact1_up, file_fact2_up]
        if all(f is None for f in csv_fontes):
            csv_fontes = glob.glob(f"*FACTURAÇÃO*{ano_corrente}_{mes_selecionado:02d}.csv")
            if not csv_fontes:
                csv_fontes = glob.glob("*FACTURAÇÃO*.csv")
            if not csv_fontes:
                csv_fontes = glob.glob("*.csv")
            
        for file_f in csv_fontes:
            if file_f is not None:
                try:
                    file_bytes = io.BytesIO(file_f.getvalue()) if hasattr(file_f, "getvalue") else open(file_f, "rb")
                    temp_df = pd.read_csv(file_bytes, sep=None, engine="python", encoding="iso-8859-1", skiprows=3, on_bad_lines='skip', header=0)
                    if not temp_df.empty and len(temp_df.columns) > 1:
                        df_fact_total = pd.concat([df_fact_total, temp_df], ignore_index=True)
                except Exception:
                    pass
        
        if df_fact_total.empty:
            st.warning("⚠️ Nenhum ficheiro CSV de faturação encontrado ou carregado.")
        else:
            idx_ofic = 0
            idx_val = 24 if df_fact_total.shape[1] > 24 else df_fact_total.shape[1] - 1
            idx_data = 28 if df_fact_total.shape[1] > 28 else df_fact_total.shape[1] - 2

            df_fact_total["Oficina_CSV"] = df_fact_total.iloc[:, idx_ofic].astype(str).str.strip()
            
            # Limpar linhas indesejadas
            invalidos = ["NAN", "", "TOTAL", "MAPA DE FACTURAÇÃO"]
            df_fact_total = df_fact_total[~df_fact_total["Oficina_CSV"].str.upper().isin(invalidos)]
            df_fact_total = df_fact_total[~df_fact_total["Oficina_CSV"].str.startswith("*")]

            # Conversão da Data
            col_data_raw = df_fact_total.iloc[:, idx_data]
            df_fact_total["Data_Parsed"] = pd.to_datetime(col_data_raw, dayfirst=True, errors="coerce")
            
            # Valores Numéricos
            col_valor_raw = df_fact_total.iloc[:, idx_val]
            if col_valor_raw.dtype == object:
                s_clean = col_valor_raw.astype(str).str.replace("€", "", regex=False).str.strip()
                s_clean = s_clean.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
                df_fact_total["Valor_Num"] = pd.to_numeric(s_clean, errors="coerce").fillna(0.0)
            else:
                df_fact_total["Valor_Num"] = pd.to_numeric(col_valor_raw, errors="coerce").fillna(0.0)

            # Agrupamento
            if not df_agrup.empty:
                df_agrup.columns = [str(c).strip() for c in df_agrup.columns]
                col_agrup_oficina = df_agrup.columns[0]
                col_agrup_destino = df_agrup.columns[3] if len(df_agrup.columns) > 3 else df_agrup.columns[1]
                dict_agrupamento = dict(zip(df_agrup[col_agrup_oficina].astype(str).str.strip().str.upper(), df_agrup[col_agrup_destino].astype(str).str.strip()))
                df_fact_total["Oficina_Agrupada"] = df_fact_total["Oficina_CSV"].str.upper().map(dict_agrupamento).fillna(df_fact_total["Oficina_CSV"])
            else:
                df_fact_total["Oficina_Agrupada"] = df_fact_total["Oficina_CSV"]

            oficinas_disponiveis_fat = sorted(df_fact_total["Oficina_Agrupada"].dropna().astype(str).unique().tolist())
            if not oficinas_disponiveis_fat:
                oficinas_disponiveis_fat = ["Geral"]

            col_sel_ofic, col_empty = st.columns([2, 2])
            with col_sel_ofic:
                oficina_selecionada_fat = st.selectbox("Selecionar Oficina (Faturação):", options=oficinas_disponiveis_fat)
                
            def limpar_str(s):
                if pd.isna(s): return ""
                s_nfkd = unicodedata.normalize('NFKD', str(s))
                return "".join([c for c in s_nfkd if not unicodedata.combining(c)]).upper().strip()

            sel_clean = limpar_str(oficina_selecionada_fat)
            
            df_oficina_fat = df_fact_total[
                df_fact_total["Oficina_Agrupada"].apply(limpar_str).str.contains(sel_clean, na=False, regex=False) |
                df_fact_total["Oficina_CSV"].apply(limpar_str).str.contains(sel_clean, na=False, regex=False)
            ]
            if df_oficina_fat.empty:
                df_oficina_fat = df_fact_total
                
            df_mes_ano = df_oficina_fat[df_oficina_fat["Data_Parsed"].dt.month == mes_selecionado]
            if not df_mes_ano.empty:
                df_oficina_fat = df_mes_ano
                
            df_diario = df_oficina_fat.groupby(df_oficina_fat["Data_Parsed"].dt.date, as_index=False)["Valor_Num"].sum()
            df_diario.rename(columns={"Valor_Num": "Facturacao_Diaria", "Data_Parsed": "Data"}, inplace=True)
            df_diario = df_diario.dropna(subset=["Data"]).sort_values("Data")
            
            inicio_mes = date(ano_corrente, mes_selecionado, 1)
            fim_mes = date(ano_corrente + 1, 1, 1) - timedelta(days=1) if mes_selecionado == 12 else date(ano_corrente, mes_selecionado + 1, 1) - timedelta(days=1)
                
            range_dias = pd.date_range(start=inicio_mes, end=fim_mes)
            df_completo_dias = pd.DataFrame({"Data": range_dias.date})
            df_completo_dias = pd.merge(df_completo_dias, df_diario, on="Data", how="left").fillna(0.0)
            df_completo_dias["Facturacao_Acumulada"] = df_completo_dias["Facturacao_Diaria"].cumsum()
            
            # Orçamento Mensal
            orcamento_mensal = 0.0
            if not df_orc.empty and len(df_orc.columns) >= 5:
                col_ofic = df_orc.columns[0]
                col_mes = df_orc.columns[3]
                col_val = df_orc.columns[4]
                nome_mes_alvo = limpar_str(meses_pt[mes_selecionado])
                
                for _, row in df_orc.iterrows():
                    ofic_row = limpar_str(row[col_ofic])
                    if sel_clean in ofic_row or ofic_row in sel_clean:
                        val_mes_cell = row[col_mes]
                        mes_encontrado = False
                        if isinstance(val_mes_cell, (pd.Timestamp, datetime, date)):
                            if val_mes_cell.month == mes_selecionado: mes_encontrado = True
                        else:
                            mes_row_str = limpar_str(val_mes_cell)
                            if nome_mes_alvo in mes_row_str or mes_row_str in nome_mes_alvo or str(mes_selecionado) in mes_row_str:
                                mes_encontrado = True
                        if mes_encontrado:
                            try:
                                v_str = str(row[col_val]).replace("€", "").replace(" ", "").replace(".", "").replace(",", ".")
                                v_num = float(v_str)
                                if v_num > 0:
                                    orcamento_mensal = v_num
                                    break
                            except:
                                pass
                                
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
                st.metric(label="ORÇAMENTO MENSAL", value=f"{orcamento_mensal:,.0f} €".replace(",", " "))
                st.metric(label="ORÇAMENTO DIÁRIO", value=f"{orcamento_diario:,.0f} €".replace(",", " "))
                st.metric(label="FATURAÇÃO À DATA", value=f"{facturacao_a_data:,.0f} €".replace(",", " "))
                st.metric(label="DESVIO ORÇAMENTO", value=f"{desvio_orcamento:,.0f} €".replace(",", " "), delta=f"{desvio_orcamento:,.0f} €".replace(",", " "))
