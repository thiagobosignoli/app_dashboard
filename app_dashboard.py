import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import os
import io
import shutil
from datetime import datetime

# 1. Configuração da Página
st.set_page_config(page_title="Dashboard Fiscal e Imobiliário", layout="wide")

st.title("📊 Sistema Analítico de Cadastro Imobiliário e IPTU")
st.write("Gestão de estoque fiscal, acompanhamento de metas, Autos de Lançamento e evolução por exercício.")

# 2. Configuração de Pasta Local / Rede Corporativa
st.sidebar.header("⚙️ Configuração de Armazenamento")
pasta_dados = st.sidebar.text_input(
    "Caminho da Pasta da Base de Dados:",
    value="base_dados",
    help="No futuro, você pode colocar um caminho de rede (ex: \\\\servidor\\pasta_fiscal)"
).strip()

pasta_historico = os.path.join(pasta_dados, "historico")
os.makedirs(pasta_dados, exist_ok=True)
os.makedirs(pasta_historico, exist_ok=True)

caminho_base_oficial = os.path.join(pasta_dados, "base_oficial.parquet")

# 3. Mapeamento das Colunas Principais (Configurável)
with st.sidebar.expander("🏷️ Mapeamento de Colunas", expanded=False):
    col_chave = st.text_input("Inscrição Imobiliária (Chave):", value="NUM_INSCRICAO").strip()
    col_ano = st.text_input("Ano de Exercício:", value="ANO_EXERCICIO").strip()
    col_lote = st.text_input("Lote Fiscal:", value="Lotefiscal").strip()
    col_area_const = st.text_input("Área Construída:", value="área construída").strip()
    col_area_corr = st.text_input("Área Corrigida:", value="área Corrigida").strip()
    col_venal = st.text_input("Valor Venal (R$):", value="valor venal").strip()
    col_iptu_ufm = st.text_input("IPTU (UFMs):", value="IPTU").strip()
    col_iptu_moeda = st.text_input("IPTU Moeda (R$):", value="IPTU Moeda").strip()
    col_al = st.text_input("Auto de Lançamento (R$):", value="AutodeLancamento").strip()

# 4. Funções Utilitárias Blindadas
def ler_arquivo_blindado(arquivo):
    nome = arquivo.name.lower()
    bytes_data = arquivo.getvalue()
    if nome.endswith(".xlsx") or nome.endswith(".xls"):
        df = pd.read_excel(io.BytesIO(bytes_data), dtype=str)
    else:
        try:
            df = pd.read_csv(io.BytesIO(bytes_data), sep=None, engine="python", encoding="utf-8-sig", dtype=str, on_bad_lines="skip")
        except UnicodeDecodeError:
            df = pd.read_csv(io.BytesIO(bytes_data), sep=None, engine="python", encoding="latin1", dtype=str, on_bad_lines="skip")
    df.columns = df.columns.str.strip()
    return df

def normalizar_chave(serie):
    return serie.astype(str).str.strip().str.replace(r"[\.\-\/]", "", regex=True).str.lstrip("0")

def converter_numero(serie):
    s = serie.astype(str).str.strip()
    s = np.where(s.str.contains(","), s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False), s)
    return pd.Series(pd.to_numeric(s, errors="coerce"), index=serie.index).fillna(0.0)

def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_numero(valor, decimais=2):
    formato = f"{{:,.{decimais}f}}"
    return formato.format(valor).replace(",", "X").replace(".", ",").replace("X", ".")

def preparar_dataframe(df):
    df = df.copy()
    if col_chave in df.columns:
        df["CHAVE_LIMPA"] = normalizar_chave(df[col_chave])
    if col_ano in df.columns:
        df[col_ano] = df[col_ano].astype(str).str.strip()
    for c in [col_area_const, col_area_corr, col_venal, col_iptu_ufm, col_iptu_moeda, col_al]:
        if c in df.columns:
            df[f"{c}_NUM"] = converter_numero(df[c])
    return df

def converter_para_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dados_Fiscais")
    return output.getvalue()

base_existe = os.path.exists(caminho_base_oficial)
# ============================================================
# CENÁRIO A: PRIMEIRA VEZ (NÃO HÁ BASE NA PASTA)
# ============================================================
if not base_existe:
    st.warning(f"⚠️ Nenhuma base oficial encontrada na pasta `{pasta_dados}`.")
    st.subheader("📥 Carga Inicial da Base de Dados")
    st.write("Envie o primeiro arquivo (CSV ou Excel) para criar o banco de dados local do dashboard.")

    arquivo_inicial = st.file_uploader("Selecione a planilha inicial (.csv ou .xlsx)", type=["csv", "xlsx", "xls"])

    if arquivo_inicial:
        if st.button("💾 Salvar como Base Oficial"):
            with st.spinner("Processando e salvando base inicial..."):
                df_ini = ler_arquivo_blindado(arquivo_inicial)
                if col_chave not in df_ini.columns:
                    st.error(f"A coluna chave '{col_chave}' não foi encontrada. Ajuste o nome no menu lateral 'Mapeamento de Colunas'.")
                else:
                    df_ini["DATA_CARGA"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    df_ini.to_parquet(caminho_base_oficial, index=False)
                    st.success("Base inicial salva com sucesso! Recarregando painel...")
                    st.rerun()

# ============================================================
# CENÁRIO B: BASE ENCONTRADA -> EXIBE O DASHBOARD COMPLETO
# ============================================================
else:
    df_raw = pd.read_parquet(caminho_base_oficial)
    df = preparar_dataframe(df_raw)

    data_ultima_carga = df["DATA_CARGA"].iloc[0] if "DATA_CARGA" in df.columns else "Desconhecida"
    st.sidebar.success(f"✅ Base Conectada\n\n**Registros:** {len(df):,}\n\n**Última Carga:** {data_ultima_carga}")

    aba_dash, aba_tempo, aba_explorar, aba_atualizar = st.tabs([
        "📈 Visão Geral (KPIs)",
        "📅 Evolução Temporal (Exercícios)",
        "🔍 Explorar e Exportar Dados",
        "🔄 Atualizar Base (Confronto / SQL)"
    ])

    # --------------------------------------------------------
    # ABA 1: VISÃO GERAL (KPIs E AUTOS DE LANÇAMENTO)
    # --------------------------------------------------------
    with aba_dash:
        df_visao = df.copy()

        # Filtro rápido de Ano de Exercício no topo da aba
        if col_ano in df_visao.columns:
            anos_disp = sorted(df_visao[col_ano].dropna().unique().tolist(), reverse=True)
            anos_sel = st.multiselect("Filtrar Visão Geral por Ano de Exercício (deixe vazio para ver tudo):", options=anos_disp)
            if anos_sel:
                df_visao = df_visao[df_visao[col_ano].isin(anos_sel)]

        st.subheader("Resumo Executivo do Cadastro e IPTU")
        k1, k2, k3, k4 = st.columns(4)

        total_insc = len(df_visao)
        soma_iptu_rs = df_visao[f"{col_iptu_moeda}_NUM"].sum() if f"{col_iptu_moeda}_NUM" in df_visao.columns else 0.0
        soma_iptu_ufm = df_visao[f"{col_iptu_ufm}_NUM"].sum() if f"{col_iptu_ufm}_NUM" in df_visao.columns else 0.0
        soma_venal = df_visao[f"{col_venal}_NUM"].sum() if f"{col_venal}_NUM" in df_visao.columns else 0.0

        k1.metric("Inscrições no Recorte", formatar_numero(total_insc, 0))
        k2.metric("Total IPTU Moeda (R$)", formatar_moeda(soma_iptu_rs))
        k3.metric("Total IPTU (UFMs)", f"{formatar_numero(soma_iptu_ufm, 2)} UFMs")
        k4.metric("Valor Venal Total", formatar_moeda(soma_venal))

        st.divider()
        st.subheader("⚖️ Autos de Lançamento (ALs) e Áreas")
        a1, a2, a3, a4 = st.columns(4)

        if f"{col_al}_NUM" in df_visao.columns:
            qtd_als = (df_visao[f"{col_al}_NUM"] > 0).sum()
            soma_als = df_visao[f"{col_al}_NUM"].sum()
        else:
            qtd_als = 0
            soma_als = 0.0

        soma_area_c = df_visao[f"{col_area_const}_NUM"].sum() if f"{col_area_const}_NUM" in df_visao.columns else 0.0
        soma_area_corr = df_visao[f"{col_area_corr}_NUM"].sum() if f"{col_area_corr}_NUM" in df_visao.columns else 0.0

        a1.metric("Quantidade de ALs", formatar_numero(qtd_als, 0))
        a2.metric("Valor Total ALs (R$)", formatar_moeda(soma_als))
        a3.metric("Área Construída Total", f"{formatar_numero(soma_area_c, 2)} m²")
        a4.metric("Área Corrigida Total", f"{formatar_numero(soma_area_corr, 2)} m²")

        # Gráfico de Distribuição por Lote Fiscal
        if col_lote in df_visao.columns and f"{col_iptu_moeda}_NUM" in df_visao.columns:
            st.divider()
            df_lote = df_visao.groupby(col_lote, as_index=False).agg(
                IPTU_Reais=(f"{col_iptu_moeda}_NUM", "sum"),
                Autos_Lancamento=(f"{col_al}_NUM", "sum") if f"{col_al}_NUM" in df_visao.columns else (col_lote, "count")
            ).sort_values("IPTU_Reais", ascending=False).head(15)

            fig_lote = px.bar(
                df_lote,
                x=col_lote,
                y="IPTU_Reais",
                title="Top 15 Lotes Fiscais por Arrecadação de IPTU (R$)",
                labels={col_lote: "Lote Fiscal", "IPTU_Reais": "IPTU Moeda (R$)"},
                text_auto=".2s"
            )
            st.plotly_chart(fig_lote, use_container_width=True)

    # --------------------------------------------------------
    # ABA 2: EVOLUÇÃO TEMPORAL (ANO_EXERCICIO)
    # --------------------------------------------------------
    with aba_tempo:
        st.subheader("📅 Análise Comparativa por Ano de Exercício")

        if col_ano not in df.columns:
            st.warning(f"A coluna '{col_ano}' não foi encontrada na base. Verifique o nome no menu lateral 'Mapeamento de Colunas'.")
        else:
            agg_dict = {col_chave: "count"}
            if f"{col_iptu_moeda}_NUM" in df.columns:
                agg_dict[f"{col_iptu_moeda}_NUM"] = "sum"
            if f"{col_iptu_ufm}_NUM" in df.columns:
                agg_dict[f"{col_iptu_ufm}_NUM"] = "sum"
            if f"{col_al}_NUM" in df.columns:
                agg_dict[f"{col_al}_NUM"] = ["sum", lambda x: (x > 0).sum()]

            df_tempo = df.groupby(col_ano).agg(agg_dict)
            # Achata as colunas multi-nível
            df_tempo.columns = ["_".join(col).strip() if isinstance(col, tuple) else col for col in df_tempo.columns]
            df_tempo = df_tempo.reset_index().sort_values(col_ano)

            # Renomeia para exibição limpa
            mapa_nomes = {
                f"{col_chave}_count": "Qtd_Inscricoes",
                f"{col_iptu_moeda}_NUM_sum": "Total_IPTU_Reais",
                f"{col_iptu_ufm}_NUM_sum": "Total_IPTU_UFMs",
                f"{col_al}_NUM_sum": "Valor_ALs_Reais",
                f"{col_al}_NUM_<lambda_0>": "Qtd_ALs"
            }
            df_tempo = df_tempo.rename(columns=mapa_nomes)

            col_g1, col_g2 = st.columns(2)
            with col_g1:
                if "Total_IPTU_Reais" in df_tempo.columns:
                    fig_iptu_tempo = px.bar(
                        df_tempo,
                        x=col_ano,
                        y="Total_IPTU_Reais",
                        title="Evolução do IPTU Moeda (R$) por Exercício",
                        labels={col_ano: "Exercício", "Total_IPTU_Reais": "IPTU (R$)"},
                        text_auto=".2s"
                    )
                    st.plotly_chart(fig_iptu_tempo, use_container_width=True)

            with col_g2:
                if "Valor_ALs_Reais" in df_tempo.columns:
                    fig_al_tempo = px.bar(
                        df_tempo,
                        x=col_ano,
                        y="Valor_ALs_Reais",
                        title="Evolução dos Autos de Lançamento (R$) por Exercício",
                        labels={col_ano: "Exercício", "Valor_ALs_Reais": "Autos de Lançamento (R$)"},
                        text_auto=".2s",
                        color_discrete_sequence=["#EF553B"]
                    )
                    st.plotly_chart(fig_al_tempo, use_container_width=True)

            st.subheader("Tabela Consolidada por Exercício")
            st.dataframe(df_tempo, use_container_width=True)

    # --------------------------------------------------------
    # ABA 3: EXPLORAR E EXPORTAR DADOS
    # --------------------------------------------------------
    with aba_explorar:
        st.subheader("Filtros e Exportação Personalizada")

        df_filtrado = df_raw.copy()

        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            if col_ano in df_filtrado.columns:
                anos_f = sorted(df_filtrado[col_ano].dropna().unique().tolist(), reverse=True)
                anos_sel_f = st.multiselect(f"Filtrar por {col_ano}:", options=anos_f, key="filt_ano_exp")
                if anos_sel_f:
                    df_filtrado = df_filtrado[df_filtrado[col_ano].isin(anos_sel_f)]
        with col_f2:
            if col_lote in df_filtrado.columns:
                lotes_disponiveis = sorted(df_filtrado[col_lote].dropna().unique().tolist())
                lotes_selecionados = st.multiselect(f"Filtrar por {col_lote}:", options=lotes_disponiveis, key="filt_lote_exp")
                if lotes_selecionados:
                    df_filtrado = df_filtrado[df_filtrado[col_lote].isin(lotes_selecionados)]
        with col_f3:
            busca_insc = st.text_input(f"Buscar {col_chave} específica:", key="filt_insc_exp").strip()
            if busca_insc and col_chave in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado[col_chave].astype(str).str.contains(busca_insc, case=False, na=False)]
        apenas_com_al = st.checkbox("Mostrar apenas inscrições com Auto de Lançamento (> 0)")
        if apenas_com_al and col_al in df_filtrado.columns:
            vals_al = converter_numero(df_filtrado[col_al])
            df_filtrado = df_filtrado[vals_al > 0]
        st.write(f"Exibindo **{len(df_filtrado):,}** registros de um total de **{len(df_raw):,}**.")
        st.dataframe(df_filtrado.head(200), use_container_width=True)
        st.divider()
        st.subheader("📥 Exportar Dados")
        exp_col1, exp_col2 = st.columns(2)
        with exp_col1:
            st.markdown("Exportar Recorte Filtrado na Tela:")
            csv_filt = df_filtrado.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button("📄 Baixar Filtrado (.csv)", data=csv_filt, file_name="recorte_fiscal.csv", mime="text/csv", key="dl_csv_filt")
            xlsx_filt = converter_para_excel(df_filtrado)
            st.download_button("📊 Baixar Filtrado (.xlsx)", data=xlsx_filt, file_name="recorte_fiscal.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_xlsx_filt")
            with exp_col2:
                st.markdown("Exportar Base Oficial Completa:")
                csv_full = df_raw.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")
                st.download_button("📄 Baixar Base Completa (.csv)", data=csv_full, file_name="base_oficial_completa.csv", mime="text/csv", key="dl_csv_full")
                xlsx_full = converter_para_excel(df_raw)
                st.download_button("📊 Baixar Base Completa (.xlsx)", data=xlsx_full, file_name="base_oficial_completa.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_xlsx_full")
# --------------------------------------------------------
# ABA 4: ATUALIZAR BASE (CONFRONTO DIÁRIO / SQL)
# --------------------------------------------------------
            with aba_atualizar:
                metodo_atualizacao = st.radio("Escolha a fonte de atualização:", ["Upload Manual (CSV / Excel)", "Extração Direta PostgreSQL (Em breve - Fase 4)"])
                if metodo_atualizacao == "Upload Manual (CSV / Excel)":
                    arquivo_novo = st.file_uploader("Selecione a nova planilha diária/mensal:", type=["csv", "xlsx", "xls"], key="up_atualizacao")
                    if arquivo_novo:
                        df_novo_raw = ler_arquivo_blindado(arquivo_novo)
                        if col_chave not in df_novo_raw.columns:
                            st.error(f"A coluna chave '{col_chave}' não existe no arquivo enviado.")
                        else:
                            chaves_antigas = set(normalizar_chave(df_raw[col_chave]).dropna())
                            chaves_novas = set(normalizar_chave(df_novo_raw[col_chave]).dropna())
                            insc_entradas = chaves_novas - chaves_antigas
                            insc_excluidas = chaves_antigas - chaves_novas
                            insc_mantidas = chaves_novas.intersection(chaves_antigas)
                            st.subheader("⚖️ Balanço de Confronto (Antes de Substituir)")
                            c1, c2, c3 = st.columns(3)
                            c1.metric("🆕 Novas Inscrições", len(insc_entradas))
                            c2.metric("❌ Inscrições Excluídas", len(insc_excluidas))
                            c3.metric("🔁 Inscrições Mantidas", len(insc_mantidas))
                            if st.button("✅ Confirmar Confronto e Atualizar Base Oficial"):
                                with st.spinner("Arquivando histórico e atualizando base oficial..."):
                                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                    caminho_backup = os.path.join(pasta_historico, f"backup_base_{timestamp}.parquet")
                                    shutil.copy2(caminho_base_oficial, caminho_backup)
                                    df_novo_raw["DATA_CARGA"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    df_novo_raw.to_parquet(caminho_base_oficial, index=False)
                                    st.success("Base atualizada com sucesso! Backup salvo na pasta histórico.")
                                    st.rerun()
                            else:
                                st.info("🔌 Configuração de Conexão PostgreSQL + Proxy (Preparado para Fase 4)")
                                col_db1, col_db2 = st.columns(2)
                                with col_db1:
                                    st.text_input("Host / URL:", value="lpostpro27", disabled=True)
                                    st.text_input("Porta / Banco:", value="5000 / smf", disabled=True)
                                    st.text_input("Usuário Banco:", placeholder="Digite seu usuário do PostgreSQL")
                                    st.text_input("Senha Banco:", type="password")
                                    with col_db2:
                                        st.text_input("Proxy Corporativo (Host:Porta):", placeholder="ex: procempa.com.br:3128")
                                        st.text_input("Usuário Proxy:", placeholder="Usuário da rede")
                                        st.text_input("Senha Proxy:", type="password")
                                        st.text_area("Query SQL de Extração:", value="SELECT * FROM sua_tabela_imobiliaria LIMIT 100;")
                                        st.button("⚡ Testar Conexão SQL (Ativaremos na Fase 4)", disabled=True)
