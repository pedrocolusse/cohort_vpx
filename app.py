# Notebook: processamento profundo de report.xlsx
# Defina abaixo os parâmetros conforme sua planilha (mude os nomes das colunas se necessário)
FILE = "report.xlsx"  # arquivo deve estar na mesma pasta do notebook
SHEET = 0  # altere se precisa ler outra sheet (0 = primeira)
DEAL_COL = "Nome do negócio"  # coluna que identifica o negócio
CANAL_COL = "Canal"  # coluna com o canal
# Coluna que contém o valor do contrato (coloque o nome exato da sua planilha aqui)
CONTRACT_COL = "Valor previsto"  # <--- ajuste este nome conforme a sua planilha
MEETING_COL = "Date entered \"Reunião realizada (SQL) (Closer - Pipeline de Vendas)\""  # coluna que contém a data em que entrou na fase 'reunião realizada'
CLOSED_COL = "Date entered \"Fechado verbalmente  (Oportunidade) (Closer - Pipeline de Vendas)\""  # coluna que contém a data em que fechou
# Strings de status que serão usadas na coluna 'Status' do resultado
STATUS_MEETING = "Reunião realizada SQL"  # string para a primeira entrada (data 1)
STATUS_CLOSED = "Fechado"  # string para a segunda entrada (data 2)
STATUS_LOST = "Perdido"  # string para a entrada de perdido
# Nome da coluna de saída de data
OUTPUT_DATE_COL = "Data do Status"
# Nome da coluna de saída para valor do contrato no DataFrame final
OUTPUT_CONTRACT_COL = "Valor do Contrato"
LOST_DATE_COL = "Date entered \"Perdidos Closer (Closer - Pipeline de Vendas)\""

# Import necessário
import streamlit as st
import pandas as pd
import numpy as np
import re
from pathlib import Path
# from annotated_text import annotated_text, annotation
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.express as px
from operator import attrgetter
from streamlit_option_menu import option_menu
import tempfile
import os

st.set_page_config(
        page_title="Sandbox Comercial VPx",
        page_icon="🏆",
        layout="wide",
    )

st.sidebar.title("Sandbox Comercial VPx")

# NOVO: Inicializar session_state para persistência de dados
if 'processed_df' not in st.session_state:
    st.session_state.processed_df = None
if 'file_uploaded' not in st.session_state:
    st.session_state.file_uploaded = False

# st.sidebar.image("logo.png", caption="S A N D B O X")

@st.cache_data
def home():
    st.subheader(
    "Bem-vindo(a) ao Sandbox VPx.")
    st.write("Este aplicativo é projetado para fornecer ferramentas úteis para o dia a dia do consultor. Explore as funcionalidades disponibilizadas:")
    st.markdown("""
    #### ***Funcionalidades:***
    - Análise de cohort, para entender o comportamento de clientes ao longo do tempo. :blue-badge[Comercial] :green-badge[Pipeline Closer]
                """)
    return

# Helper para converter strings monetárias em float
def parse_money(value):
    """Tenta converter vários formatos comuns de dinheiro em float.
    Exemplos aceitos: 'R$ 1.234,56', '1234.56', '$1,234.56', '1.234', '1234,56'
    Retorna np.nan se não for possível converter.
    """
    if value is None:
        return np.nan
    if isinstance(value, (int, float, np.floating, np.integer)):
        return float(value)
    s = str(value).strip()
    if s == "":
        return np.nan
    # remover símbolos de moeda e espaços
    s = re.sub(r"[^0-9,\.\-]", "", s)
    if s == "" or s == "-":
        return np.nan
    # Se tiver tanto '.' quanto ',' assumimos que um é separador de milhares
    # regra heurística:
    # - se houver '.' e ',' e '.' aparece antes de ',', então '.' é milhares e ',' é decimal
    # - se houver '.' e ',' e ',' aparece antes de '.', então ',' é milhares e '.' é decimal
    try:
        if "." in s and "," in s:
            if s.rfind('.') < s.rfind(','):
                # exemplo: 1.234,56 -> remove '.' e troca ',' por '.'
                s = s.replace('.', '').replace(',', '.')
            else:
                # exemplo: 1,234.56 -> remove ','
                s = s.replace(',', '')
        elif "," in s:
            # Se só houver vírgula, pode ser decimal ou milhares
            # heurística: se houver mais de 1 dígito após a vírgula e <=2, tratar como decimal
            parts = s.split(',')
            if len(parts[-1]) in (1,2):
                s = s.replace(',', '.')
            else:
                # ex: '1,234' -> remover ','
                s = s.replace(',', '')
        else:
            # só ponto ou nenhum separador: remover espaços já feito
            s = s
        return float(s)
    except Exception:
        return np.nan

# Função principal de processamento
def process_report(file_path: str,
                   sheet=0,
                   deal_col: str = DEAL_COL,
                   canal_col: str = CANAL_COL,
                   contract_col: str = CONTRACT_COL,
                   meeting_col: str = MEETING_COL,
                   closed_col: str = CLOSED_COL,
                   status_meeting: str = STATUS_MEETING,
                   status_closed: str = STATUS_CLOSED,
                   status_lost: str = STATUS_LOST,
                   output_date_col: str = OUTPUT_DATE_COL,
                   output_contract_col: str = OUTPUT_CONTRACT_COL,
                   lost_date_col: str = LOST_DATE_COL):
    """
    Processa o arquivo e retorna um DataFrame com colunas:
    Nome do negócio | Canal | Valor do Contrato | Status | Data do Status

    Regras aplicadas:
    - Para cada negócio, busca a primeira data (mais antiga) encontrada na coluna `meeting_col` (se existir)
      e, separadamente, a primeira data encontrada na coluna `closed_col`.
    - Gera até 2 linhas por negócio: primeira com a data de `meeting_col` e segunda com `closed_col` quando existirem
      E APENAS se a primeira (meeting_col) existir. Se houver somente fechamento sem reunião, o negócio é excluído.
    - Se apenas `meeting_col` existir, produz somente a primeira linha.
    - `canal_col` e `contract_col` são extraídos como o primeiro valor não-nulo encontrado para o negócio.
    - A coluna de saída de data está formatada como string 'DD/MM/YYYY'.
    - A coluna de contrato será convertida em numérica (float). Se não for possível, ficará NaN e será substituída por 0 no final.
    """
    p = Path(file_path)
    if not p.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")

    # Leitura
    df = pd.read_excel(p, sheet_name=sheet, engine="openpyxl")
    print(f"Arquivo lido com {len(df):,} linhas. Colunas detectadas: {list(df.columns)}")

    # Conferir colunas pedidas (deal e canal obrigatórios)
    missing = [c for c in (deal_col, canal_col) if c not in df.columns]
    if missing:
        raise KeyError(f"Colunas obrigatórias ausentes no arquivo: {missing}")

    # Normalizar colunas de data para datetime (não string ainda)
    for c in (meeting_col, closed_col, lost_date_col):
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    # Se existir a coluna de contrato, vamos tentar convertê-la para numérico desde já
    if contract_col in df.columns:
        df['_contract_numeric_parsed'] = df[contract_col].apply(parse_money)
    else:
        df['_contract_numeric_parsed'] = np.nan

    # Agrupar por negócio e escolher a primeira (mais antiga) ocorrência das datas, canal e contrato
    result_rows = []
    grouped = df.groupby(deal_col, dropna=False)
    for deal, g in grouped:
        # canal: primeiro valor não-nulo encontrado
        canal_val = None
        if canal_col in g.columns:
            non_null = g[canal_col].dropna()
            if len(non_null) > 0:
                canal_val = non_null.iloc[0]

        # contrato: primeiro valor numérico não-nulo encontrado (se coluna existir)
        contract_val = np.nan
        if '_contract_numeric_parsed' in g.columns:
            c_non_null = g['_contract_numeric_parsed'].dropna()
            if len(c_non_null) > 0:
                contract_val = float(c_non_null.iloc[0])

        # obter a primeira (menor) data para meeting_col, closed_col e lost_date_col
        meet_date = pd.NaT
        close_date = pd.NaT
        lost_date = pd.NaT
        if meeting_col in g.columns:
            meet_dates = pd.to_datetime(g[meeting_col], errors="coerce").dropna()
            if not meet_dates.empty:
                meet_date = meet_dates.min()
        if closed_col in g.columns:
            close_dates = pd.to_datetime(g[closed_col], errors="coerce").dropna()
            if not close_dates.empty:
                close_date = close_dates.min()
        if lost_date_col in g.columns:
            lost_dates = pd.to_datetime(g[lost_date_col], errors="coerce").dropna()
            if not lost_dates.empty:
                lost_date = lost_dates.min()

        # Regras de inclusão de linhas
        # 1) Se existir meet_date, cria primeira entrada com status_meeting
        if pd.notnull(meet_date):
            result_rows.append({deal_col: deal, canal_col: canal_val, contract_col: contract_val, "Status": status_meeting, output_date_col: meet_date})
            # 2) Se houver close_date E existir meet_date, cria segunda entrada com status_closed
            if pd.notnull(close_date):
                result_rows.append({deal_col: deal, canal_col: canal_val, contract_col: contract_val, "Status": status_closed, output_date_col: close_date})
        # 3) Se NÃO existir meet_date mas existir close_date: NÃO INCLUI (regra solicitada)
        
        # 4) Se existir lost_date, cria entrada com status_lost, APENAS SE JÁ NÃO EXISTIR UM "Fechado" PARA O MESMO NEGÓCIO
        if pd.notnull(lost_date):
            has_closed = any(row for row in result_rows if row[deal_col] == deal and row["Status"] == status_closed)
            meeting_before_lost = pd.notnull(meet_date) and (lost_date > meet_date)
            if not has_closed and meeting_before_lost:
                result_rows.append({deal_col: deal, canal_col: canal_val, contract_col: contract_val, "Status": status_lost, output_date_col: lost_date})

    out = pd.DataFrame(result_rows)

    # Reordenar colunas para o formato pedido e formatar a coluna de data como DD/MM/YYYY
    if not out.empty:
        # garantir coluna do canal com nome fixo 'Canal' se o usuário tiver outra coluna
        rename_map = {}
        if canal_col != "Canal":
            rename_map[canal_col] = "Canal"
        # renomear a coluna de contrato para o nome de saída desejado
        if contract_col and contract_col != output_contract_col:
            rename_map[contract_col] = output_contract_col
        if rename_map:
            out = out.rename(columns=rename_map)

        # assegurar as colunas finais na ordem: deal, Canal, Valor do Contrato, Status, Data do Status
        final_cols = [deal_col, "Canal"]
        if output_contract_col in out.columns:
            final_cols.append(output_contract_col)
        final_cols.extend(["Status", output_date_col])
        out = out[final_cols]

        # garantir que a coluna de contrato seja numérica e substituir NaN por 0
        if output_contract_col in out.columns:
            out[output_contract_col] = pd.to_numeric(out[output_contract_col], errors='coerce').fillna(0.0)

        # converter para datetime e formatar como string DD/MM/YYYY
        out[output_date_col] = pd.to_datetime(out[output_date_col], errors="coerce").dt.strftime('%d/%m/%Y')
        out = out.sort_values([deal_col, output_date_col], na_position='last').reset_index(drop=True)
    else:
        # criar dataframe vazio com colunas corretas
        cols = [deal_col, "Canal", "Status", output_date_col]
        # incluir coluna de contrato mesmo vazia
        if contract_col:
            cols.insert(2, output_contract_col)
        out = pd.DataFrame(columns=cols)

    # remover coluna auxiliar caso exista no df de origem
    if '_contract_numeric_parsed' in df.columns:
        df = df.drop(columns=['_contract_numeric_parsed'])

    return df, out

def cohort():
    st.sidebar.markdown("""Esta seção permitirá que você realize uma análise por safra em seus dados.Suba sua base de dados em Excel (.xlsx) e selecione os canais e datas para a análise de cohort.
    """)
    
    # NOVO: File uploader com key única
    uploaded_file = st.sidebar.file_uploader("Escolha um arquivo Excel (.xlsx)", type="xlsx", key="file_uploader")

    # NOVO: Processar arquivo quando for carregado
    if uploaded_file is not None and not st.session_state.file_uploaded:
        try:
            # Salva o arquivo enviado em um temporário e processa com process_report
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                tmp.write(uploaded_file.getbuffer())
                tmp_path = tmp.name

            # process_report retorna (raw_df, final_df)
            _, processed_df = process_report(tmp_path, sheet=0)

            # remover arquivo temporário
            try:
                os.remove(tmp_path)
            except Exception:
                pass

            # NOVO: Salvar no session_state
            st.session_state.processed_df = processed_df
            st.session_state.file_uploaded = True
            st.sidebar.success("✅ Arquivo processado com sucesso!")

        except Exception as e:
            st.error(f"Erro ao processar arquivo: {str(e)}")
            return

    # NOVO: Botão para limpar/resetar os dados
    if st.session_state.file_uploaded:
        if st.sidebar.button("🗑️ Limpar dados e carregar novo arquivo"):
            st.session_state.processed_df = None
            st.session_state.file_uploaded = False
            st.rerun()

    # NOVO: Usar dados do session_state
    if st.session_state.processed_df is not None:
        try:
            df_original = st.session_state.processed_df.copy()
            
            # Exibir as primeiras linhas do DataFrame original (já processado)
            if st.checkbox("**Pré-visualização dos dados originais (processados):**"):
                st.dataframe(df_original.head(500))

            # Faz uma cópia para trabalhar
            df = df_original.copy()

            # Filtrar por canais via multiseleção
            if 'Canal' in df.columns:
                channels = df['Canal'].dropna().unique().tolist()
            else:
                channels = []
            selected_channels = st.sidebar.multiselect("Filtrar canais (Canal):", options=channels, default=channels)
            if selected_channels:
                df = df[df['Canal'].isin(selected_channels)]

            # Colunas fixas: data = 'Data do Status', cliente = 'Nome do negócio'
            date_column = 'Data do Status'
            client_column = 'Nome do negócio'

            # Converter coluna de data para datetime com formato padronizado
            df[date_column] = pd.to_datetime(df[date_column], format='%d/%m/%Y', errors='coerce')

            # Filtro por intervalo de datas (opcional)
            date_filter = st.sidebar.date_input("Selecione o intervalo de datas:", [], format="DD/MM/YYYY")
            date_filter = [pd.to_datetime(date) for date in date_filter] if date_filter else None
            if date_filter and len(date_filter) == 2:
                df = df[(df[date_column] >= date_filter[0]) & (df[date_column] <= date_filter[1])]

            # Selecionar período para a coorte
            periodo = st.sidebar.selectbox("Selecione o período para a coorte:", ["M", "Y", "Q"], index=0,
                                          format_func=lambda x: "Mensal" if x == "M" else "Anual" if x == "Y" else "Trimestral")

            # Processamento dos dados para análise de coorte
            df['ano_cohort'] = df[date_column].dt.to_period(periodo)
            df['cohort'] = df.groupby(client_column)[date_column].transform('min').dt.to_period(periodo)
            
            # CORREÇÃO: Identificar TODOS os cohorts do dataframe original
            all_cohorts = sorted(df['cohort'].dropna().unique())
            
            # Separar dados por status
            df_fechados = df[df['Status'] == 'Fechado']
            df_perdidos = df[df['Status'] == 'Perdido']
            
            # Calcular o tamanho total da coorte (para a coluna branca) - TODOS os cohorts
            cohort_sizes = df.groupby('cohort')[client_column].nunique()
            # CORREÇÃO: Reindexar para incluir todos os cohorts, preenchendo com 0 os vazios
            cohort_sizes = cohort_sizes.reindex(all_cohorts, fill_value=0)
            
            # NOVO: Calcular métricas adicionais por cohort
            # Quantidade de fechados por cohort
            cohort_closed = df_fechados.groupby('cohort')[client_column].nunique()
            cohort_closed = cohort_closed.reindex(all_cohorts, fill_value=0)

            #Quantidade de perdidos por cohort
            cohort_lost = df_perdidos.groupby('cohort')[client_column].nunique()
            cohort_lost = cohort_lost.reindex(all_cohorts, fill_value=0)
            
            # Quantidade restante (não fechados)
            cohort_remaining = cohort_sizes - cohort_closed - cohort_lost
            cohort_remaining = cohort_remaining.clip(lower=0)  # evitar negativos
            
            # Percentual de conversão
            cohort_conversion = (cohort_closed / cohort_sizes * 100).fillna(0)
            
            # Agrupar considerando apenas fechados
            df_grouped = df_fechados.groupby(['cohort', 'ano_cohort']).agg(n_customers1=(client_column, 'nunique')).reset_index(drop=False)
            df_grouped['period_number'] = (df_grouped.ano_cohort - df_grouped.cohort).apply(attrgetter('n'))

            # Criar pivot table com os dados de fechados
            cohort_pivot = df_grouped.pivot_table(index='cohort', columns='period_number', values='n_customers1')
            
            # CORREÇÃO: Reindexar o pivot para incluir todos os cohorts
            cohort_pivot = cohort_pivot.reindex(all_cohorts)
            
            # Usar o tamanho total da coorte para calcular a retenção
            cohort_size = cohort_sizes  # Agora já contém todos os cohorts
            retention_matrix = cohort_pivot.divide(cohort_size, axis=0)

            # Adicionar toggle para alternar entre porcentagem e valores absolutos
            show_percentage = st.sidebar.toggle('Mostrar valores em porcentagem', value=False)
            
            # Criar gráfico interativo com Plotly - NOVO: passar métricas adicionais
            if show_percentage:
                create_cohort_heatmap_plotly(retention_matrix, cohort_size, cohort_conversion, cohort_closed, cohort_remaining, cohort_lost, is_percentage=True)
            else:
                create_cohort_heatmap_plotly(cohort_pivot, cohort_size, cohort_conversion, cohort_closed, cohort_remaining, cohort_lost, is_percentage=False)

        except Exception as e:
            st.error(f"Erro na análise: {str(e)}")
    else:
        st.info("Aguardando o upload do Excel do Hubspot.")

def create_cohort_heatmap_plotly(matrix, cohort_size, cohort_conversion, cohort_closed, cohort_remaining, cohort_lost, is_percentage=True):
    # Preparar dados para o gráfico
    cohort_labels = matrix.index.astype(str).tolist()
    # keep the original column values (these are the actual period numbers, may be non-contiguous)
    period_cols = list(matrix.columns)
    # Ensure numeric period columns (they should be ints), attempt to coerce
    try:
        period_ints = [int(x) for x in period_cols]
    except Exception:
        # fallback: use positional indices if coercion fails
        period_ints = list(range(len(period_cols)))

    # Ensure we always show up to period 5 (user request). Compute the full range to include
    # any existing max period or 5, whichever is larger.
    if period_ints:
        min_p = min(period_ints)
        max_p = max(period_ints)
    else:
        min_p = 0
        max_p = 0
    desired_max = max(max_p, 5)
    full_periods = list(range(min_p, desired_max + 1))

    # Reindex the matrix to include missing period columns (filled with NaN). This ensures
    # the heatmap has a column for period 5 even if there is no data for it.
    matrix = matrix.reindex(columns=full_periods)

    # Update period_cols and labels to the full range
    period_cols = full_periods
    # string labels for axis display
    period_labels = [str(col) for col in period_cols]
    
    # NOVO: Criar labels para o sales cycle (em dias)
    sales_cycle_labels = []
    for period in period_cols:
        if period == 0:
            sales_cycle_labels.append("<30")
        elif period == 1:
            sales_cycle_labels.append("30-60")
        else:
            # Para período 2+: calcular intervalos de 30 em 30 dias
            inicio = 30 * period
            fim = 30 * (period + 1)
            sales_cycle_labels.append(f"{inicio}-{fim}")
    
    # NOVO: Criar uma figura com subplots expandida: 5 colunas
    # Coluna 1: Tamanho cohort
    # Coluna 2: Matriz principal
    # Coluna 3: % Conversão
    # Coluna 4: Qtd Fechados
    # Coluna 5: Qtd Restante
    fig = make_subplots(
        rows=1, cols=6,
        column_widths=[0.05, 0.65, 0.08, 0.06, 0.06, 0.06],  # Ajustar larguras
        shared_yaxes=True,
        horizontal_spacing=0.005,
    )
    
    # Adicionar heatmap para o tamanho da coorte (coluna 1)
    cohort_size_values = cohort_size.values.reshape(-1, 1)
    
    fig.add_trace(
        go.Heatmap(
            z=cohort_size_values,
            x=['Safra'],
            y=cohort_labels,
            text=cohort_size_values,
            texttemplate="%{text}",
            textfont={"size": 16},
            colorscale=[[0, 'white'], [1, 'white']],
            showscale=False,
            hoverinfo='text',
            hovertext=[[f'Tamanho da cohort: {int(val)}' for val in cohort_size_values]],
        ),
        row=1, col=1
    )
    
    # Adicionar heatmap para a matriz de retenção (coluna 2)
    mask = np.isnan(matrix.values)
    
    # Criar textos para exibição
    text_matrix = []
    hover_matrix = []
    
    z_values = matrix.copy()
    
    for i in range(matrix.shape[0]):
        text_row = []
        hover_row = []
        for j in range(matrix.shape[1]):
            if not np.isnan(matrix.iloc[i, j]):
                period_value = period_cols[j]
                if is_percentage:
                    text_row.append(f"{matrix.iloc[i, j]:.0%}")
                    hover_row.append(f"<b>Safra: {cohort_labels[i]}</b><br>Período: {period_value}<br>Taxa de conversão: {matrix.iloc[i, j]:.1%}")
                else:
                    text_row.append(f"{int(matrix.iloc[i, j])}")
                    hover_row.append(f"<b>Safra: {cohort_labels[i]}</b><br>Período: {period_value}<br>Quantidade: {int(matrix.iloc[i, j])}")
            else:
                text_row.append("")
                hover_row.append("")
        text_matrix.append(text_row)
        hover_matrix.append(hover_row)
    
    fig.add_trace(
        go.Heatmap(
            z=z_values,
            x=period_labels,
            y=cohort_labels,
            text=text_matrix,
            texttemplate="%{text}",
            textfont={"size": 18},
            hoverinfo='text',
            hovertext=hover_matrix,
            colorscale='RdYlGn',
            colorbar=dict(
                title='Taxa de Conversão' if is_percentage else 'Quantidade de Deals',
                tickformat='.0%' if is_percentage else 'd',
                ticks='outside'
            ),
            zmin=0,
            zmax=1 if is_percentage else z_values.max().max(),
            connectgaps=False,
            hoverongaps=False,
            showscale=True,
            zauto=False,
            zmid=0.5,
        ),
        row=1, col=2
    )
    
    # NOVO: CORREÇÃO - Adicionar coluna 3 - % Conversão
    conversion_values = cohort_conversion.values.reshape(-1, 1)
    # Criar texto formatado manualmente
    conversion_text = [[f"{val:.1f}%" for val in conversion_values.flatten()]]
    conversion_text = list(zip(*conversion_text))  # Transpor para formato correto
    
    fig.add_trace(
        go.Heatmap(
            z=conversion_values,
            x=['% Conv.'],
            y=cohort_labels,
            text=conversion_text,
            texttemplate="%{text}",
            textfont={"size": 16},
            colorscale=[[0, 'white'], [1, 'white']],
            showscale=False,
            hoverinfo='text',
            hovertext=[[f'<b>Conversão: {val:.1f}%</b>' for val in conversion_values.flatten()]],
        ),
        row=1, col=3
    )
    
    # NOVO: CORREÇÃO - Adicionar coluna 4 - Qtd Fechados
    closed_values = cohort_closed.values.reshape(-1, 1)
    # Criar texto formatado manualmente
    closed_text = [[str(int(val)) for val in closed_values.flatten()]]
    closed_text = list(zip(*closed_text))  # Transpor
    
    fig.add_trace(
        go.Heatmap(
            z=closed_values,
            x=['Fechados'],
            y=cohort_labels,
            text=closed_text,
            texttemplate="%{text}",
            textfont={"size": 16},
            colorscale=[[0, 'white'], [1, 'white']],
            showscale=False,
            hoverinfo='text',
            hovertext=[[f'Fechados: {int(val)}' for val in closed_values.flatten()]],
        ),
        row=1, col=4
    )
    
    # NOVO: CORREÇÃO - Adicionar coluna 5 - Qtd Restante
    lost_values = cohort_lost.values.reshape(-1, 1)
    # Criar texto formatado manualmente
    lost_text = [[str(int(val)) for val in lost_values.flatten()]]
    lost_text = list(zip(*lost_text))  # Transpor
    
    fig.add_trace(
        go.Heatmap(
            z=lost_values,
            x=['Perdidos'],
            y=cohort_labels,
            text=lost_text,
            texttemplate="%{text}",
            textfont={"size": 16},
            colorscale=[[0, 'white'], [1, 'white']],
            showscale=False,
            hoverinfo='text',
            hovertext=[[f'Restante: {int(val)}' for val in lost_values.flatten()]],
        ),
        row=1, col=5
    )

    # NOVO: CORREÇÃO - Adicionar coluna 5 - Qtd Restante
    remaining_values = cohort_remaining.values.reshape(-1, 1)
    # Criar texto formatado manualmente
    remaining_text = [[str(int(val)) for val in remaining_values.flatten()]]
    remaining_text = list(zip(*remaining_text))  # Transpor
    
    fig.add_trace(
        go.Heatmap(
            z=remaining_values,
            x=['Restante'],
            y=cohort_labels,
            text=remaining_text,
            texttemplate="%{text}",
            textfont={"size": 16},
            colorscale=[[0, 'white'], [1, 'white']],
            showscale=False,
            hoverinfo='text',
            hovertext=[[f'Restante: {int(val)}' for val in remaining_values.flatten()]],
        ),
        row=1, col=6
    )

    # Atualizar layout
    fig.update_layout(
        title_text="Análise de Cohort - Taxa de Retenção",
        height=700,
        width=1400,  # Aumentar largura para acomodar novas colunas
        yaxis=dict(
            title='',
            autorange='reversed',
        ),
        xaxis2=dict(
            title='# períodos',
            tickmode='array',
            tickvals=list(range(len(period_labels))),
            ticktext=period_labels,
            side='top',
        ),
    )
    
    # Adicionar segundo eixo X com sales cycle
    fig.update_xaxes(
        title='# sales cycle (dias)',
        tickmode='array',
        tickvals=list(range(len(sales_cycle_labels))),
        ticktext=sales_cycle_labels,
        side='bottom',
        row=1, col=2
    )

    # Ajustar margens
    fig.update_layout(margin=dict(l=50, r=50, t=100, b=100))
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=False)
    
    # Exibir o gráfico no Streamlit
    st.plotly_chart(fig, use_container_width=True)
    
    return

def sideBar():
    with st.sidebar:
        selected=option_menu(
            menu_title="Menu Principal",
            options=["Home", "Cohort Analysis"],
            icons = ["house", "graph-up-arrow"],
            menu_icon="cast",
            default_index=0,
            styles={
        "container": {},
        "icon": {"color": "white", "font-size": "15px"}, 
        "nav-link": {"font-family": "sans-serif", "font-size": "15px", "text-align": "left", "margin":"0px"},
        "nav-link-selected": {"background-color": "#D70248"},
    }
        )
        
    if selected=="Home":
        home()
    if selected=="Cohort Analysis":
        st.subheader(f"{selected} :blue-badge[Comercial] :green-badge[Pipeline Closer]")
        cohort()

sideBar()