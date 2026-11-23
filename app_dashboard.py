import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import os
import io
import shutil
from datetime import datetime

# 1. Configuração da Página
st.set_page_config(page_title="Painel DRI - IPTU e Cadastro", layout="wide")

st.title("📊 Painel Analítico - IPTU e Cadastro Imobiliário")
st.caption("Lançamento, evolução por exercício, estoque imobiliário e Autos de Lançamento.")

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

# 3. Mapeamento das Colunas Principais (Atualizado com os nomes reais da imagem)
with st.sidebar.expander("🏷️ Mapeamento de Colunas", expanded=False):
    col_chave = st.text_input("Inscrição Imobiliária (Chave):", value="NUM_INSCRICAO").strip()
    col_ano = st.text_input("Ano de Exercício:", value="ANO_EXERCICIO").strip()
    col_tipo_lanc = st.text_input("Tipo de Lançamento:", value="TIPO_LANCAMENTO").strip()
    col_bairro = st.text_input("Bairro:", value="NOME_BAIRRO").strip()
    col_lote = st.text_input("Lote Fiscal:", value="NUM_BLOCO").strip()
    col_tipologia = st.text_input("Tipologia (Casa/Apto/Terreno):", value="DES_FINALIDADE").strip()
    col_uso = st.text_input("Uso (Residencial/Comercial):", value="DES_USO").strip()
    col_area_const = st.text_input("Área Construída:", value="MTR_AREA_REAL").strip()
    col_area_corr = st.text_input("Área Corrigida:", value="MTR_AREA_CORRIGIDA").strip()
    col_venal = st.text_input("Valor Venal (R$):", value="VLR_VENAL_IMOVEL_TRIB").strip()
    col_iptu_ufm = st.text_input("IPTU (UFMs):", value="VLR_IPTU_UFM").strip()
    col_iptu_moeda = st.text_input("IPTU Moeda (R$):", value="VLR_IPTU_MOEDA").strip()
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

def formatar_moeda_curta(valor):
    if abs(valor) >= 1_000_000_000:
        return f"R$ {valor/1_000_000_000:,.2f} bi".replace(",", "X").replace(".", ",").replace("X", ".")
    elif abs(valor) >= 1_000_000:
        return f"R$ {valor/1_000_000:,.1f} mi".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_numero(valor, decimais=0):
    formato = f"{{:,.{decimais}f}}"
    return formato.format(valor).replace(",", "X").replace(".", ",").replace("X", ".")

def calcular_delta(atual, anterior):
    if anterior == 0 or pd.isna(anterior):
        return None
    diff_pct = ((atual - anterior) / anterior) * 100
    sinal = "+" if diff_pct > 0 else ""
    return f"{sinal}{diff_pct:,.1f}% vs. ano ant.".replace(".", ",")

def preparar_dataframe(df):
    df = df.copy()
    if col_chave in df.columns:
        df["CHAVE_LIMPA"] = normalizar_chave(df[col_chave])
    if col_ano in df.columns:
        df[col_ano] = df[col_ano].astype(str).str.strip()
    if col_tipo_lanc in df.columns:
        df[col_tipo_lanc] = df[col_tipo_lanc].fillna("Não Informado").astype(str).str.strip()
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

    data_ultima_carga = df["DATA_CARGA"].iloc if "DATA_CARGA" in df.columns else "Desconhecida"
    st.sidebar.success(f"✅ Base Conectada\n\n**Registros:** {len(df):,}\n\n**Última Carga:** {data_ultima_carga}")

    aba_dash, aba_explorar, aba_atualizar = st.tabs([
        "📈 Painel Executivo (IPTU & Estoque)",
        "🔍 Explorar e Exportar Dados",
        "🔄 Atualizar Base (Confronto / SQL)"
    ])

    # --------------------------------------------------------
    # ABA 1: PAINEL EXECUTIVO (ESTILO PROTÓTIPO DRI)
    # --------------------------------------------------------
    with aba_dash:
        # Linha 1 de Filtros Globais (Exercício e Tipo de Lançamento em destaque)
        f_top1, f_top2 = st.columns([1, 2.5])
        
        ano_selecionado = "Todos os anos"
        with f_top1:
            if col_ano in df.columns:
                anos_disp = sorted(df[col_ano].dropna().unique().tolist(), reverse=True)
                opcoes_ano = ["Todos os anos"] + anos_disp
                ano_selecionado = st.selectbox("EXERCÍCIO:", options=opcoes_ano, index=1 if len(anos_disp) > 0 else 0)

        with f_top2:
            tipos_lanc_sel = []
            if col_tipo_lanc in df.columns:
                tipos_disp = sorted(df[col_tipo_lanc].dropna().unique().tolist())
                tipos_lanc_sel = st.multiselect("TIPO DE LANÇAMENTO:", options=tipos_disp, placeholder="Todos os tipos de lançamento")

        # Linha 2 de Filtros Estruturais (Bairro, Lote Fiscal e Uso)
        f_col1, f_col2, f_col3 = st.columns(3)
        with f_col1:
            bairros_sel = []
            if col_bairro in df.columns:
                bairros_disp = sorted(df[col_bairro].dropna().unique().tolist())
                bairros_sel = st.multiselect("BAIRRO:", options=bairros_disp, placeholder="Todos os bairros")

        with f_col2:
            lotes_sel_dash = []
            if col_lote in df.columns:
                lotes_disp = sorted(df[col_lote].dropna().unique().tolist())
                lotes_sel_dash = st.multiselect("LOTE FISCAL (BLOCO):", options=lotes_disp, placeholder="Todos os lotes")

        with f_col3:
            usos_sel = []
            if col_uso in df.columns:
                usos_disp = sorted(df[col_uso].dropna().unique().tolist())
                usos_sel = st.multiselect("USO:", options=usos_disp, placeholder="Todos os usos")

        # Aplica filtros estruturais antes da separação anual
        df_base_filtro = df.copy()
        if tipos_lanc_sel and col_tipo_lanc in df_base_filtro.columns:
            df_base_filtro = df_base_filtro[df_base_filtro[col_tipo_lanc].isin(tipos_lanc_sel)]
        if bairros_sel and col_bairro in df_base_filtro.columns:
            df_base_filtro = df_base_filtro[df_base_filtro[col_bairro].isin(bairros_sel)]
        if lotes_sel_dash and col_lote in df_base_filtro.columns:
            df_base_filtro = df_base_filtro[df_base_filtro[col_lote].isin(lotes_sel_dash)]
        if usos_sel and col_uso in df_base_filtro.columns:
            df_base_filtro = df_base_filtro[df_base_filtro[col_uso].isin(usos_sel)]

        # Separa recorte atual e recorte do ano anterior
        if col_ano in df_base_filtro.columns and ano_selecionado != "Todos os anos":
            df_atual = df_base_filtro[df_base_filtro[col_ano] == ano_selecionado]
            try:
                ano_anterior = str(int(ano_selecionado) - 1)
                df_ant = df_base_filtro[df_base_filtro[col_ano] == ano_anterior]
            except ValueError:
                df_ant = pd.DataFrame()
        else:
            df_atual = df_base_filtro.copy()
            df_ant = pd.DataFrame()

        # Cálculo dos 6 KPIs Principais
        insc_atual = len(df_atual)
        insc_ant = len(df_ant)

        iptu_rs_atual = df_atual[f"{col_iptu_moeda}_NUM"].sum() if f"{col_iptu_moeda}_NUM" in df_atual.columns else 0.0
        iptu_rs_ant = df_ant[f"{col_iptu_moeda}_NUM"].sum() if (not df_ant.empty and f"{col_iptu_moeda}_NUM" in df_ant.columns) else 0.0

        iptu_ufm_atual = df_atual[f"{col_iptu_ufm}_NUM"].sum() if f"{col_iptu_ufm}_NUM" in df_atual.columns else 0.0
        iptu_ufm_ant = df_ant[f"{col_iptu_ufm}_NUM"].sum() if (not df_ant.empty and f"{col_iptu_ufm}_NUM" in df_ant.columns) else 0.0

        al_rs_atual = df_atual[f"{col_al}_NUM"].sum() if f"{col_al}_NUM" in df_atual.columns else 0.0
        al_qtd_atual = int((df_atual[f"{col_al}_NUM"] > 0).sum()) if f"{col_al}_NUM" in df_atual.columns else 0
        al_rs_ant = df_ant[f"{col_al}_NUM"].sum() if (not df_ant.empty and f"{col_al}_NUM" in df_ant.columns) else 0.0

        venal_atual = df_atual[f"{col_venal}_NUM"].sum() if f"{col_venal}_NUM" in df_atual.columns else 0.0
        venal_ant = df_ant[f"{col_venal}_NUM"].sum() if (not df_ant.empty and f"{col_venal}_NUM" in df_ant.columns) else 0.0

        ticket_atual = (iptu_rs_atual / insc_atual) if insc_atual > 0 else 0.0
        ticket_ant = (iptu_rs_ant / insc_ant) if insc_ant > 0 else 0.0

        st.divider()
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.metric("Inscrições lançadas", formatar_numero(insc_atual), delta=calcular_delta(insc_atual, insc_ant))
        k2.metric("Valor lançado IPTU", formatar_moeda_curta(iptu_rs_atual), delta=calcular_delta(iptu_rs_atual, iptu_rs_ant))
        k3.metric("Lançado em UFMs", f"{formatar_numero(iptu_ufm_atual, 1)}", delta=calcular_delta(iptu_ufm_atual, iptu_ufm_ant))
        k4.metric(f"Autos Lanç. ({al_qtd_atual} ALs)", formatar_moeda_curta(al_rs_atual), delta=calcular_delta(al_rs_atual, al_rs_ant))
        k5.metric("Valor Venal Total", formatar_moeda_curta(venal_atual), delta=calcular_delta(venal_atual, venal_ant))
        k6.metric("Ticket Médio IPTU", formatar_moeda(ticket_atual), delta=calcular_delta(ticket_atual, ticket_ant))

        st.divider()

        # Grade de Gráficos
        g_esq, g_dir = st.columns([1.15, 1])
        with g_esq:
            st.markdown("**Evolução Temporal por Tipo de Lançamento**")
            
            if col_tipo_lanc in df_base_filtro.columns and col_ano in df_base_filtro.columns:
                metrica_tipo = st.radio(
                    "Visualizar evolução por:",
                    ["Valor Lançado IPTU (R$)", "Quantidade de Inscrições"],
                    horizontal=True,
                    key="radio_graf_tipo"
                )

                df_graf_tempo = df_base_filtro.copy()
                df_graf_tempo[col_tipo_lanc] = df_graf_tempo[col_tipo_lanc].astype(str).str.strip()
                df_graf_tempo[col_ano] = df_graf_tempo[col_ano].astype(str).str.strip()

                top_tipos = df_graf_tempo[col_tipo_lanc].value_counts().head(8).index
                df_graf_tempo[col_tipo_lanc] = np.where(
                    df_graf_tempo[col_tipo_lanc].isin(top_tipos),
                    df_graf_tempo[col_tipo_lanc],
                    "Outros Tipos"
                )

                if metrica_tipo == "Valor Lançado IPTU (R$)" and f"{col_iptu_moeda}_NUM" in df_graf_tempo.columns:
                    df_agrup = df_graf_tempo.groupby([col_ano, col_tipo_lanc], as_index=False)[f"{col_iptu_moeda}_NUM"].sum()
                    col_y_val = f"{col_iptu_moeda}_NUM"
                    label_y = "IPTU Moeda (R$)"
                else:
                    df_agrup = df_graf_tempo.groupby([col_ano, col_tipo_lanc], as_index=False)[col_chave].count()
                    col_y_val = col_chave
                    label_y = "Quantidade de Inscrições"

                df_agrup = df_agrup.sort_values(col_ano)

                # Criação da variável fig_tempo_tipo
                fig_tempo_tipo = px.bar(
                    df_agrup,
                    x=col_ano,
                    y=col_y_val,
                    color=col_tipo_lanc,
                    barmode="group",
                    labels={col_ano: "Exercício", col_y_val: label_y, col_tipo_lanc: "Tipo de Lançamento"},
                    text_auto=".2s"
                )

                # Trava a largura das colunas para não ficarem grossas quando houver poucos anos
                fig_tempo_tipo.update_traces(
                    textposition="outside",
                    cliponaxis=False,
                    width=0.22
                )
                fig_tempo_tipo.update_xaxes(type="category", title="Ano de Exercício")
                fig_tempo_tipo.update_yaxes(title=label_y, showgrid=True, gridcolor="#f0f2f6")
                
                # Controla o espaçamento entre os anos e posiciona a legenda na base
                fig_tempo_tipo.update_layout(
                    height=450,
                    bargap=0.35,
                    bargroupgap=0.15,
                    margin=dict(l=10, r=10, t=25, b=10),
                    plot_bgcolor="rgba(0,0,0,0)",
                    legend=dict(
                        orientation="h",
                        yanchor="top",
                        y=-0.2,
                        xanchor="center",
                        x=0.5,
                        title=""
                    )
                )
                st.plotly_chart(fig_tempo_tipo, use_container_width=True)
            else:
                st.info(f"Mapeie as colunas '{col_tipo_lanc}' e '{col_ano}' na barra lateral para visualizar este gráfico.")
            with g_dir:
                st.markdown("**Comparação anual do lançamento (IPTU R$)**")
                if col_ano in df_base_filtro.columns and f"{col_iptu_moeda}_NUM" in df_base_filtro.columns:
                df_anual = df_base_filtro.groupby(col_ano, as_index=False)[f"{col_iptu_moeda}_NUM"].sum().sort_values(col_ano)
                cores = ["#1f5aa6" if str(a) == str(ano_selecionado) else "#b0b8c4" for a in df_anual[col_ano]]

                fig_anual = go.Figure(data=[go.Bar(
                    x=df_anual[col_ano],
                    y=df_anual[f"{col_iptu_moeda}_NUM"],
                    marker_color=cores,
                    text=[formatar_moeda_curta(v) for v in df_anual[f"{col_iptu_moeda}_NUM"]],
                    textposition="outside",
                    width=0.35,
                    cliponaxis=False
                )])
                fig_anual.update_xaxes(type="category")
                fig_anual.update_yaxes(showgrid=True, gridcolor="#f0f2f6", title="IPTU (R$)")
                fig_anual.update_layout(height=240, bargap=0.45, margin=dict(l=10, r=10, t=25, b=10), plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_anual, use_container_width=True)
            st.markdown("**Imóveis edificados e não edificados**")
            if f"{col_area_const}_NUM" in df_atual.columns and insc_atual > 0:
                edificados = int((df_atual[f"{col_area_const}_NUM"] > 0).sum())
                nao_edificados = insc_atual - edificados
                pct_edif = (edificados / insc_atual) * 100
                pct_nao_edif = (nao_edificados / insc_atual) * 100

                df_edif = pd.DataFrame({
                    "Tipo": ["Não edificados (territorial)", "Edificados"],
                    "Quantidade": [nao_edificados, edificados],
                    "Texto": [f"{formatar_numero(nao_edificados)} ({pct_nao_edif:.1f}%)", f"{formatar_numero(edificados)} ({pct_edif:.1f}%)"]
                })

                fig_edif = px.bar(df_edif, x="Quantidade", y="Tipo", orientation="h", text="Texto")
                fig_edif.update_traces(marker_color="#1f77b4", textposition="outside")
                fig_edif.update_layout(height=180, margin=dict(l=10, r=60, t=10, b=10), xaxis_visible=False, yaxis_title="")
                st.plotly_chart(fig_edif, use_container_width=True)

        # Resumo por Tipo de Lançamento (se houver a coluna na base)
        if col_tipo_lanc in df_atual.columns:
            st.divider()
            st.markdown("**Resumo por Tipo de Lançamento**")
            agg_tipo = {col_chave: "count"}
            if f"{col_iptu_moeda}_NUM" in df_atual.columns:
                agg_tipo[f"{col_iptu_moeda}_NUM"] = "sum"
            if f"{col_iptu_ufm}_NUM" in df_atual.columns:
                agg_tipo[f"{col_iptu_ufm}_NUM"] = "sum"
            if f"{col_al}_NUM" in df_atual.columns:
                agg_tipo[f"{col_al}_NUM"] = "sum"

            df_tab_tipo = df_atual.groupby(col_tipo_lanc, as_index=False).agg(agg_tipo)
            df_tab_tipo = df_tab_tipo.rename(columns={
                col_chave: "Inscrições",
                f"{col_iptu_moeda}_NUM": "IPTU Lançado (R$)",
                f"{col_iptu_ufm}_NUM": "IPTU (UFMs)",
                f"{col_al}_NUM": "Autos de Lançamento (R$)"
            }).sort_values("Inscrições", ascending=False)
            st.dataframe(df_tab_tipo, use_container_width=True, hide_index=True)

        # Tabela Consolidada Inferior (Finalidade + Uso)
        st.divider()
        cols_grupo_tab = [c for c in [col_tipologia, col_uso] if c in df_atual.columns]
        if not cols_grupo_tab and col_lote in df_atual.columns:
            cols_grupo_tab = [col_lote]

        if cols_grupo_tab:
            st.markdown(f"**Lançamento e consolidação por {' / '.join(cols_grupo_tab)}**")
            agg_resumo = {col_chave: "count"}
            if f"{col_iptu_moeda}_NUM" in df_atual.columns:
                agg_resumo[f"{col_iptu_moeda}_NUM"] = "sum"
            if f"{col_iptu_ufm}_NUM" in df_atual.columns:
                agg_resumo[f"{col_iptu_ufm}_NUM"] = "sum"
            if f"{col_al}_NUM" in df_atual.columns:
                agg_resumo[f"{col_al}_NUM"] = "sum"
            if f"{col_venal}_NUM" in df_atual.columns:
                agg_resumo[f"{col_venal}_NUM"] = "sum"

            df_tab_resumo = df_atual.groupby(cols_grupo_tab, as_index=False).agg(agg_resumo)
            df_tab_resumo = df_tab_resumo.rename(columns={
                col_chave: "Inscrições",
                f"{col_iptu_moeda}_NUM": "IPTU Lançado (R$)",
                f"{col_iptu_ufm}_NUM": "IPTU (UFMs)",
                f"{col_al}_NUM": "Autos de Lançamento (R$)",
                f"{col_venal}_NUM": "Valor Venal (R$)"
            })
            if "IPTU Lançado (R$)" in df_tab_resumo.columns:
                df_tab_resumo = df_tab_resumo.sort_values("IPTU Lançado (R$)", ascending=False)

            st.dataframe(df_tab_resumo, use_container_width=True, hide_index=True)

    # --------------------------------------------------------
    # ABA 2: EXPLORAR E EXPORTAR DADOS
    # --------------------------------------------------------
    with aba_explorar:
        st.subheader("Filtros e Exportação Personalizada")
        df_filtrado = df_raw.copy()

        col_f1, col_f2, col_f3, col_f4 = st.columns(4)
        with col_f1:
            if col_ano in df_filtrado.columns:
                anos_f = sorted(df_filtrado[col_ano].dropna().unique().tolist(), reverse=True)
                anos_sel_f = st.multiselect(f"Filtrar por {col_ano}:", options=anos_f, key="filt_ano_exp")
                if anos_sel_f:
                    df_filtrado = df_filtrado[df_filtrado[col_ano].isin(anos_sel_f)]
        with col_f2:
            if col_tipo_lanc in df_filtrado.columns:
                tipos_f = sorted(df_filtrado[col_tipo_lanc].fillna("Não Informado").unique().tolist())
                tipos_sel_f = st.multiselect(f"Filtrar por {col_tipo_lanc}:", options=tipos_f, key="filt_tipo_exp")
                if tipos_sel_f:
                    df_filtrado = df_filtrado[df_filtrado[col_tipo_lanc].fillna("Não Informado").isin(tipos_sel_f)]
        with col_f3:
            if col_lote in df_filtrado.columns:
                lotes_disponiveis = sorted(df_filtrado[col_lote].dropna().unique().tolist())
                lotes_selecionados = st.multiselect(f"Filtrar por {col_lote}:", options=lotes_disponiveis, key="filt_lote_exp")
                if lotes_selecionados:
                    df_filtrado = df_filtrado[df_filtrado[col_lote].isin(lotes_selecionados)]
        with col_f4:
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
            st.markdown("**Exportar Recorte Filtrado na Tela:**")
            csv_filt = df_filtrado.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button("📄 Baixar Filtrado (.csv)", data=csv_filt, file_name="recorte_fiscal.csv", mime="text/csv", key="dl_csv_filt")
            xlsx_filt = converter_para_excel(df_filtrado)
            st.download_button("📊 Baixar Filtrado (.xlsx)", data=xlsx_filt, file_name="recorte_fiscal.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_xlsx_filt")

        with exp_col2:
            st.markdown("**Exportar Base Oficial Completa:**")
            csv_full = df_raw.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button("📄 Baixar Base Completa (.csv)", data=csv_full, file_name="base_oficial_completa.csv", mime="text/csv", key="dl_csv_full")
            xlsx_full = converter_para_excel(df_raw)
            st.download_button("📊 Baixar Base Completa (.xlsx)", data=xlsx_full, file_name="base_oficial_completa.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_xlsx_full")

    # --------------------------------------------------------
    # ABA 3: ATUALIZAR BASE (CONFRONTO DIÁRIO / SQL)
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
            st.info("🔌 **Configuração de Conexão PostgreSQL + Proxy (Preparado para Fase 4)**")
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
