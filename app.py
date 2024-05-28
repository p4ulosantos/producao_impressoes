import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import random
import datetime
import locale

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# LER O ARQUIVO

def ler_excel(arquivo):
    if arquivo:
        try:
            df = pd.read_excel(arquivo, sheet_name="Dados")
            return df
        except Exception as e:
            st.error(f"Erro ao ler o arquivo: {e}")
    return None

# CRIAR TABELAS

def criar_tabelas(df):
    tabela_pb = df.groupby(["Período", "CC", "Máquina"]).sum().reset_index()[["Período", "CC", "Máquina", "Produção P&B", "Valor Produção P&B (R$)"]]
    tabela_pb["Valor Produção P&B (R$)"] = tabela_pb["Valor Produção P&B (R$)"].apply(lambda x: locale.currency(x, grouping=True))
    
    tabela_color = df[df["Produção Color"] != 0].groupby(["Período", "CC", "Máquina"]).sum().reset_index()[["Período", "CC", "Máquina", "Produção Color", "Valor Produção Color (R$)"]]
    tabela_color["Valor Produção Color (R$)"] = tabela_color["Valor Produção Color (R$)"].apply(lambda x: locale.currency(x, grouping=True))
    
    tabela_pb_color = df.groupby(["Período", "CC", "Máquina"]).sum().reset_index()[["Período", "CC", "Máquina", "Produção P&B + Color", "Produção P&B", "Produção Color", "Total (R$)"]]
    tabela_pb_color["Total (R$)"] = tabela_pb_color["Total (R$)"].apply(lambda x: locale.currency(x, grouping=True))

    return tabela_pb, tabela_color, tabela_pb_color

# EXIBIR A TABELA

def exibir_tabela(tabela):
    st.write(tabela)

# EXIBIR O GRAFICO / MAQUINA

def exibir_grafico_barra(tabela, coluna_producao, coluna_valor, nome_coluna):
    cc_selecionados = st.multiselect("Selecione os centros de custo:", tabela["CC"].unique(), default=tabela["CC"].unique())
    tabela_filtrada = tabela[tabela["CC"].isin(cc_selecionados)]
    maquinas = tabela_filtrada["Máquina"].unique()
    maquinas_selecionadas = st.multiselect("Selecione as máquinas:", maquinas, default=maquinas)
    tabela_filtrada = tabela_filtrada[tabela_filtrada["Máquina"].isin(maquinas_selecionadas)]
    colors = [f"#{random.randint(0, 255):02x}{random.randint(0, 255):02x}{random.randint(0, 255):02x}" for _ in range(len(maquinas_selecionadas))]
    
    fig = go.Figure()
    for i, maquina in enumerate(maquinas_selecionadas):
        tabela_maquina = tabela_filtrada[tabela_filtrada["Máquina"] == maquina]
        fig.add_trace(go.Bar(x=tabela_maquina["Período"], y=tabela_maquina[coluna_producao], name=maquina, marker_color=colors[i]))
        fig.update_traces(text=tabela_filtrada[coluna_producao], textposition="auto")
        if len(maquinas_selecionadas) == 1:
            fig.add_trace(go.Scatter(x=tabela_filtrada["Período"], y=tabela_filtrada[coluna_valor], name="Produção", line=dict(color="green")))

    fig.update_layout(xaxis_tickfont_size=12, yaxis_tickfont_size=10, height=600, width=700, title=nome_coluna)
    st.plotly_chart(fig)
    st.write(tabela_filtrada)

    # Calcular total produzido no período selecionado
    total_producao = tabela_filtrada[coluna_producao].sum()
    st.write(f"Total Produzido ({coluna_producao}): {total_producao}")

# EXIBIR O GRAFICO / CC

def exibir_grafico_producao_pb_color(tabela):
    ano_selecionado = st.selectbox("Selecione o ano:", sorted(tabela["Período"].dt.year.unique()), index=len(tabela["Período"].dt.year.unique())-1)
    tabela_ano = tabela[tabela["Período"].dt.year == ano_selecionado]
    cc_selecionados = st.multiselect("Selecione os Centros de Custo:", tabela_ano["CC"].unique(), default=tabela_ano["CC"].unique())
    tabela_filtrada = tabela_ano[tabela_ano["CC"].isin(cc_selecionados)]
    producao_pb_color_por_periodo = tabela_filtrada.groupby("Período")[["Produção P&B", "Produção Color", "Produção P&B + Color"]].sum().reset_index()
    total_pb = producao_pb_color_por_periodo["Produção P&B"].sum()
    total_color = producao_pb_color_por_periodo["Produção Color"].sum()
    total_pb_color = producao_pb_color_por_periodo["Produção P&B + Color"].sum()
    st.write(f"Total de Produção P&B para o(s) Centro(s) de Custo selecionado(s): {total_pb}")
    st.write(f"Total de Produção Color para o(s) Centro(s) de Custo selecionado(s): {total_color}")
    st.write(f"Total de Produção P&B + Color para o(s) Centro(s) de Custo selecionado(s): {total_pb_color}")
    fig = go.Figure()
    fig.add_trace(go.Bar(x=producao_pb_color_por_periodo["Período"], y=producao_pb_color_por_periodo["Produção P&B"], name='Produção P&B', text=producao_pb_color_por_periodo["Produção P&B"], textposition='outside'))
    fig.add_trace(go.Bar(x=producao_pb_color_por_periodo["Período"], y=producao_pb_color_por_periodo["Produção Color"], name='Produção Color', text=producao_pb_color_por_periodo["Produção Color"], textposition='outside'))
    fig.update_layout(barmode='group', xaxis_title='Período', yaxis_title='Produção', title=f'Produção P&B e Color por Período {cc_selecionados}')
    fig.update_xaxes(tickmode='array', tickvals=producao_pb_color_por_periodo["Período"], ticktext=[datetime.datetime.strftime(date, '%B %y') for date in producao_pb_color_por_periodo["Período"]], tickangle=45, tickfont=dict(color='black'))
    st.plotly_chart(fig)
    st.write(producao_pb_color_por_periodo)

def main():
    st.title("Análise de Produção de Impressões - Mensal")
    st.subheader("Tecnologia da Informação - Monte Tabor")
    st.write("---")
    arquivo = st.file_uploader("Faça o upload do Relatório para análise:", type="xlsx")
    if arquivo is not None:
        df = ler_excel(arquivo)
        if df is not None:
            st.success(f"{arquivo.name.replace('.xlsx', '')} carregado com sucesso!")
            tabela_pb, tabela_color, tabela_pb_color = criar_tabelas(df)
            st.subheader("Selecione o tipo de produção a ser analisado: ")
            tabela_selecionada = st.selectbox("", ("Produção P&B", "Produção Color", "Produção P&B e Color"))
            st.write("---")
            if tabela_selecionada == "Produção P&B":
                if "Produção P&B" not in tabela_pb.columns:
                    st.error("Não existem dados de Produção P&B nesta tabela.")
                else:
                    st.header("Produção P&B")
                    exibir_grafico_barra(tabela_pb, "Produção P&B", "Produção P&B", "Produção P&B")
            elif tabela_selecionada == "Produção Color":
                if "Produção Color" not in tabela_color.columns:
                    st.error("Não existem dados de Produção Color nesta tabela.")
                else:
                    st.header("Produção Color")
                    exibir_grafico_barra(tabela_color, "Produção Color",  "Produção Color", "Produção Color")
            elif tabela_selecionada == "Produção P&B e Color":
                if "Produção P&B" not in tabela_pb_color.columns or "Produção Color" not in tabela_pb_color.columns:
                    st.error("Não existem dados de Produção P&B e Color nesta tabela.")
                else:
                    st.header("Produção P&B e Color")
                    exibir_grafico_barra(tabela_pb_color,  "Produção P&B + Color", "Produção P&B + Color", "Produção P&B + Color")
            st.header("Produção P&B e Color por Período e CC")
            exibir_grafico_producao_pb_color(tabela_pb_color)

if __name__ == "__main__":
    main()
