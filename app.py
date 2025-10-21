# Notebook: processamento profundo de report.xlsx
# Defina abaixo os parâmetros conforme sua planilha (mude os nomes das colunas se necessário)
FILE = "report.xlsx"  # arquivo deve estar na mesma pasta do notebook
SHEET = 0  # altere se precisa ler outra sheet (0 = primeira)
DEAL_COL = "Nome do negócio"  # coluna que identifica o negócio
CANAL_COL = "Canal"  # coluna com o canal
FATURAMENTO_COL = "Faturamento Anual"  # coluna com faturamento anual
# Coluna que contém o valor do contrato (coloque o nome exato da sua planilha aqui)
CONTRACT_COL = "Valor"  # <--- ajuste este nome conforme a sua planilha
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
OUTPUT_FATURAMENTO_COL = "Faturamento Anual"
LOST_DATE_COL = "Date entered \"Perdidos Closer (Closer - Pipeline de Vendas)\""
# NOVO: Colunas para verificar se negócio voltou após perdido
NEGOTIATION_COL = "Date entered \"Em negociação (Oportunidade) (Closer - Pipeline de Vendas)\""
CLOSING_COL = "Date entered \"Fechamento (Oportunidade) (Closer - Pipeline de Vendas)\""


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


# Helper para formatar período como texto legível
def format_period(period, period_type='M'):
    """
    Converte um período pandas para formato legível
    M (mensal): 2025-01 -> Jan/2025
    Q (trimestral): 2025Q1 -> Q1/2025
    Y (anual): 2025 -> 2025
    """
    if pd.isna(period):
        return ""
    
    period_str = str(period)
    
    if period_type == 'M':
        # Formato: 2025-01 -> Jan/2025
        meses = {
            '01': 'Jan', '02': 'Fev', '03': 'Mar', '04': 'Abr',
            '05': 'Mai', '06': 'Jun', '07': 'Jul', '08': 'Ago',
            '09': 'Set', '10': 'Out', '11': 'Nov', '12': 'Dez'
        }
        try:
            year, month = period_str.split('-')
            return f"{meses.get(month, month)}/{year}"
        except:
            return period_str
    
    elif period_type == 'Q':
        # Formato: 2025Q1 -> Q1/2025
        try:
            return period_str.replace('Q', 'T') + '/20' if 'Q' in period_str else period_str
        except:
            return period_str
    
    elif period_type == 'Y':
        # Formato: 2025 -> 2025
        return period_str
    
    return period_str


# Função principal de processamento
def process_report(file_path: str,
                   sheet=0,
                   deal_col: str = DEAL_COL,
                   canal_col: str = CANAL_COL,
                   contract_col: str = CONTRACT_COL,
                   faturamento_col: str = FATURAMENTO_COL,
                   meeting_col: str = MEETING_COL,
                   closed_col: str = CLOSED_COL,
                   status_meeting: str = STATUS_MEETING,
                   status_closed: str = STATUS_CLOSED,
                   status_lost: str = STATUS_LOST,
                   output_date_col: str = OUTPUT_DATE_COL,
                   output_contract_col: str = OUTPUT_CONTRACT_COL,
                   output_faturamento_col: str = OUTPUT_FATURAMENTO_COL,
                   lost_date_col: str = LOST_DATE_COL,
                   negotiation_col: str = NEGOTIATION_COL,
                   closing_col: str = CLOSING_COL):
    """
    Processa o arquivo e retorna um DataFrame com colunas:
    Nome do negócio | Canal | Programa | Faturamento Anual | Valor do Contrato | Status | Data do Status
    """
    p = Path(file_path)
    if not p.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")


    # Leitura
    df = pd.read_excel(p, sheet_name=sheet, engine="openpyxl")


    # Conferir colunas pedidas (deal e canal obrigatórios)
    missing = [c for c in (deal_col, canal_col) if c not in df.columns]
    if missing:
        raise KeyError(f"Colunas obrigatórias ausentes no arquivo: {missing}")


    # Normalizar colunas de data para datetime (não string ainda)
    for c in (meeting_col, closed_col, lost_date_col, negotiation_col, closing_col):
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")


    # Se existir a coluna de contrato, vamos tentar convertê-la para numérico desde já
    if contract_col in df.columns:
        df['_contract_numeric_parsed'] = df[contract_col].apply(parse_money)
    else:
        df['_contract_numeric_parsed'] = np.nan

    # Se existir a coluna de faturamento, vamos garantir que esteja em formato numérico
    if faturamento_col in df.columns:
        df['_faturamento_numeric_parsed'] = pd.to_numeric(df[faturamento_col], errors='coerce')
    else:
        df['_faturamento_numeric_parsed'] = np.nan


    # Agrupar por negócio e escolher a primeira (mais antiga) ocorrência das datas, canal, programa, faturamento e contrato
    result_rows = []
    grouped = df.groupby(deal_col, dropna=False)
    for deal, g in grouped:
        # canal: primeiro valor não-nulo encontrado
        canal_val = None
        if canal_col in g.columns:
            non_null = g[canal_col].dropna()
            if len(non_null) > 0:
                canal_val = non_null.iloc[0]

        # NOVO: programa: primeiro valor não-nulo encontrado, preenche com "Vazio" se não encontrar
        programa_val = "Vazio"  # Valor padrão
        if 'Programa' in g.columns:
            non_null = g['Programa'].dropna()
            if len(non_null) > 0:
                val = str(non_null.iloc[0]).strip()
                if val and val.lower() not in ['', 'nan', 'none']:
                    programa_val = val

        # faturamento: primeiro valor numérico não-nulo encontrado (se coluna existir)
        faturamento_val = np.nan
        if '_faturamento_numeric_parsed' in g.columns:
            f_non_null = g['_faturamento_numeric_parsed'].dropna()
            if len(f_non_null) > 0:
                faturamento_val = float(f_non_null.iloc[0])

        # contrato: primeiro valor numérico não-nulo encontrado (se coluna existir)
        contract_val = np.nan
        if '_contract_numeric_parsed' in g.columns:
            c_non_null = g['_contract_numeric_parsed'].dropna()
            if len(c_non_null) > 0:
                contract_val = float(c_non_null.iloc[0])


        # obter a primeira (menor) data para meeting_col e as datas máximas/mais recentes para as demais
        meet_date = pd.NaT
        close_date = pd.NaT
        lost_date = pd.NaT
        negotiation_date = pd.NaT
        closing_date = pd.NaT
        
        if meeting_col in g.columns:
            meet_dates = pd.to_datetime(g[meeting_col], errors="coerce").dropna()
            if not meet_dates.empty:
                meet_date = meet_dates.min()
                
        if closed_col in g.columns:
            close_dates = pd.to_datetime(g[closed_col], errors="coerce").dropna()
            if not close_dates.empty:
                close_date = close_dates.max()  # Pegar a mais recente
                
        if lost_date_col in g.columns:
            lost_dates = pd.to_datetime(g[lost_date_col], errors="coerce").dropna()
            if not lost_dates.empty:
                lost_date = lost_dates.max()  # Pegar a mais recente
        
        if negotiation_col in g.columns:
            negotiation_dates = pd.to_datetime(g[negotiation_col], errors="coerce").dropna()
            if not negotiation_dates.empty:
                negotiation_date = negotiation_dates.max()  # Pega a mais recente
                
        if closing_col in g.columns:
            closing_dates = pd.to_datetime(g[closing_col], errors="coerce").dropna()
            if not closing_dates.empty:
                closing_date = closing_dates.max()  # Pega a mais recente


        # NOVO: Verificar se CLOSED_COL ou LOST_DATE_COL são iguais a MEETING_COL
        closed_equals_meeting = False
        lost_equals_meeting = False
        
        if pd.notnull(meet_date) and pd.notnull(close_date):
            # Comparar apenas a data (sem hora) para evitar problemas de precisão
            if meet_date.date() == close_date.date():
                closed_equals_meeting = True
        
        if pd.notnull(meet_date) and pd.notnull(lost_date):
            # Comparar apenas a data (sem hora) para evitar problemas de precisão
            if meet_date.date() == lost_date.date():
                lost_equals_meeting = True


        # NOVO: Determinar o status final baseado na data mais recente entre TODAS as fases
        # Coletar todas as datas disponíveis com seus respectivos status
        status_dates = []
        
        if pd.notnull(close_date):
            status_dates.append(('CLOSED', close_date))
        
        if pd.notnull(lost_date):
            status_dates.append(('LOST', lost_date))
        
        if pd.notnull(negotiation_date):
            status_dates.append(('NEGOTIATION', negotiation_date))
        
        if pd.notnull(closing_date):
            status_dates.append(('CLOSING', closing_date))
        
        # Determinar qual é o status mais recente
        is_really_closed = False
        is_really_lost = False
        
        # CASO ESPECIAL: Se fechado ou perdido são iguais à reunião, têm prioridade
        if closed_equals_meeting:
            is_really_closed = True
            is_really_lost = False
        elif lost_equals_meeting:
            is_really_closed = False
            is_really_lost = True
        # CASO NORMAL: Verificar data mais recente
        elif status_dates:
            # Ordenar por data (mais recente primeiro)
            status_dates.sort(key=lambda x: x[1], reverse=True)
            most_recent_status, most_recent_date = status_dates[0]
            
            # Definir o status final baseado na fase mais recente
            if most_recent_status == 'CLOSED':
                is_really_closed = True
                is_really_lost = False
            elif most_recent_status == 'LOST':
                is_really_closed = False
                is_really_lost = True
            elif most_recent_status in ['NEGOTIATION', 'CLOSING']:
                # Se a fase mais recente é negociação ou fechamento (mas não fechado verbal),
                # o negócio está em aberto (restante)
                is_really_closed = False
                is_really_lost = False


        # Regras de inclusão de linhas
        # 1) Se existir meet_date, cria primeira entrada com status_meeting
        if pd.notnull(meet_date):
            result_rows.append({
                deal_col: deal, 
                canal_col: canal_val,
                'Programa': programa_val,  # NOVO: Adicionar programa (já com "Vazio" se necessário)
                faturamento_col: faturamento_val, 
                contract_col: contract_val, 
                "Status": status_meeting, 
                output_date_col: meet_date
            })
            
            # 2) Se está realmente fechado, cria segunda entrada com status_closed
            if is_really_closed:
                result_rows.append({
                    deal_col: deal, 
                    canal_col: canal_val,
                    'Programa': programa_val,  # NOVO: Adicionar programa (já com "Vazio" se necessário)
                    faturamento_col: faturamento_val, 
                    contract_col: contract_val, 
                    "Status": status_closed, 
                    output_date_col: close_date
                })
            
            # 3) Se está realmente perdido, cria entrada com status_lost
            # APENAS SE existir reunião antes ou igual ao perdido
            if is_really_lost:
                meeting_before_or_equal_lost = pd.notnull(meet_date) and (lost_date >= meet_date)
                if meeting_before_or_equal_lost:
                    result_rows.append({
                        deal_col: deal, 
                        canal_col: canal_val,
                        'Programa': programa_val,  # NOVO: Adicionar programa (já com "Vazio" se necessário)
                        faturamento_col: faturamento_val, 
                        contract_col: contract_val, 
                        "Status": status_lost, 
                        output_date_col: lost_date
                    })


    out = pd.DataFrame(result_rows)


    # Reordenar colunas para o formato pedido e formatar a coluna de data como DD/MM/YYYY
    if not out.empty:
        # garantir coluna do canal com nome fixo 'Canal' se o usuário tiver outra coluna
        rename_map = {}
        if canal_col != "Canal":
            rename_map[canal_col] = "Canal"
        # renomear a coluna de faturamento para o nome de saída desejado
        if faturamento_col and faturamento_col != output_faturamento_col:
            rename_map[faturamento_col] = output_faturamento_col
        # renomear a coluna de contrato para o nome de saída desejado
        if contract_col and contract_col != output_contract_col:
            rename_map[contract_col] = output_contract_col
        if rename_map:
            out = out.rename(columns=rename_map)


        # MODIFICADO: assegurar as colunas finais na ordem: deal, Canal, Programa, Faturamento Anual, Valor do Contrato, Status, Data do Status
        final_cols = [deal_col, "Canal", "Programa"]
        if output_faturamento_col in out.columns:
            final_cols.append(output_faturamento_col)
        if output_contract_col in out.columns:
            final_cols.append(output_contract_col)
        final_cols.extend(["Status", output_date_col])
        
        # Filtrar apenas colunas que existem
        final_cols = [col for col in final_cols if col in out.columns]
        out = out[final_cols]

        # garantir que a coluna de faturamento seja numérica e substituir NaN por 0
        if output_faturamento_col in out.columns:
            out[output_faturamento_col] = pd.to_numeric(out[output_faturamento_col], errors='coerce').fillna(0.0)

        # garantir que a coluna de contrato seja numérica e substituir NaN por 0
        if output_contract_col in out.columns:
            out[output_contract_col] = pd.to_numeric(out[output_contract_col], errors='coerce').fillna(0.0)

        # NOVO: Garantir que Programa não tenha valores nulos/vazios
        if 'Programa' in out.columns:
            out['Programa'] = out['Programa'].fillna('Vazio')
            out['Programa'] = out['Programa'].replace('', 'Vazio')

        # converter para datetime e formatar como string DD/MM/YYYY
        out[output_date_col] = pd.to_datetime(out[output_date_col], errors="coerce").dt.strftime('%d/%m/%Y')
        out = out.sort_values([deal_col, output_date_col], na_position='last').reset_index(drop=True)
    else:
        # criar dataframe vazio com colunas corretas
        cols = [deal_col, "Canal", "Programa", "Status", output_date_col]
        # incluir coluna de faturamento mesmo vazia
        if faturamento_col:
            cols.insert(3, output_faturamento_col)
        # incluir coluna de contrato mesmo vazia
        if contract_col:
            cols.insert(4, output_contract_col)
        out = pd.DataFrame(columns=cols)


    # remover colunas auxiliares caso existam no df de origem
    if '_contract_numeric_parsed' in df.columns:
        df = df.drop(columns=['_contract_numeric_parsed'])
    if '_faturamento_numeric_parsed' in df.columns:
        df = df.drop(columns=['_faturamento_numeric_parsed'])


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

            # NOVO: Filtrar por Programa via multiselect
            if 'Programa' in df.columns:
                programas = df['Programa'].dropna().unique().tolist()
                if programas:
                    selected_programas = st.sidebar.multiselect("Filtrar programas (Programa):", options=programas, default=programas)
                    if selected_programas:
                        df = df[df['Programa'].isin(selected_programas)]

            # NOVO: Filtrar por Faturamento Anual via select_slider
            if 'Faturamento Anual' in df.columns:
                faturamentos = df['Faturamento Anual'].dropna()
                if len(faturamentos) > 0:
                    min_fat = float(faturamentos.min())
                    max_fat = float(faturamentos.max())
                    
                    # Criar faixas de faturamento para o slider
                    faixas = [0, 1_000_000, 5_000_000, 10_000_000, 50_000_000, 100_000_000, 500_000_000, max_fat]
                    # Garantir que as faixas estejam ordenadas e únicas
                    faixas = sorted(list(set([f for f in faixas if min_fat <= f <= max_fat])))
                    
                    # Adicionar min_fat se não estiver na lista
                    if min_fat not in faixas:
                        faixas = [min_fat] + faixas
                    
                    if len(faixas) > 1:
                        selected_fat_range = st.sidebar.select_slider(
                            "Filtrar por Faturamento Anual (R$):",
                            options=faixas,
                            value=(min_fat, max_fat),
                            format_func=lambda x: f"R$ {x:,.0f}".replace(',', '.')
                        )
                        
                        # Aplicar filtro
                        df = df[(df['Faturamento Anual'] >= selected_fat_range[0]) & 
                               (df['Faturamento Anual'] <= selected_fat_range[1])]


            # Colunas fixas: data = 'Data do Status', cliente = 'Nome do negócio'
            date_column = 'Data do Status'
            client_column = 'Nome do negócio'


            # Converter coluna de data para datetime com formato padronizado
            df[date_column] = pd.to_datetime(df[date_column], format='%d/%m/%Y', errors='coerce')


            # CORRIGIDO: Filtro por intervalo de datas (APENAS para safras - Reunião realizada SQL)
            date_filter = st.sidebar.date_input("Selecione o intervalo de datas da safra:", [], format="DD/MM/YYYY")
            date_filter = [pd.to_datetime(date) for date in date_filter] if date_filter else None
            
            if date_filter and len(date_filter) == 2:
                # Identificar os clientes (negócios) que têm reunião realizada no intervalo selecionado
                df_meetings = df[df['Status'] == 'Reunião realizada SQL'].copy()
                clients_in_date_range = df_meetings[
                    (df_meetings[date_column] >= date_filter[0]) & 
                    (df_meetings[date_column] <= date_filter[1])
                ][client_column].unique()
                
                # Filtrar o DataFrame completo mantendo TODOS os status desses clientes
                df = df[df[client_column].isin(clients_in_date_range)]


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
            
            # NOVO: Calcular tempo médio de fechamento por cohort
            # Para cada cohort, calcular a média de dias entre reunião realizada e fechamento
            cohort_avg_days = {}
            for cohort in all_cohorts:
                # Pegar todos os negócios dessa cohort
                cohort_deals = df[df['cohort'] == cohort][client_column].unique()
                days_list = []
                
                for deal in cohort_deals:
                    # Pegar as datas de reunião e fechamento para esse negócio
                    deal_data = df[df[client_column] == deal].sort_values(date_column)
                    
                    # Verificar se tem status de reunião e fechamento
                    meeting_rows = deal_data[deal_data['Status'] == 'Reunião realizada SQL']
                    closed_rows = deal_data[deal_data['Status'] == 'Fechado']
                    
                    if not meeting_rows.empty and not closed_rows.empty:
                        meeting_date = meeting_rows.iloc[0][date_column]
                        closed_date = closed_rows.iloc[0][date_column]
                        
                        if pd.notnull(meeting_date) and pd.notnull(closed_date):
                            days_diff = (closed_date - meeting_date).days
                            if days_diff >= 0:  # apenas valores positivos
                                days_list.append(days_diff)
                
                # Calcular média
                if days_list:
                    cohort_avg_days[cohort] = np.mean(days_list)
                else:
                    cohort_avg_days[cohort] = 0
            
            # Converter para Series e reindexar
            cohort_avg_days = pd.Series(cohort_avg_days).reindex(all_cohorts, fill_value=0)
            
            # NOVO: Calcular Ticket Médio e Ticket Acumulado por cohort
            cohort_avg_ticket = {}
            cohort_total_ticket = {}
            
            for cohort in all_cohorts:
                # Pegar apenas os negócios fechados dessa cohort
                cohort_closed_deals = df_fechados[df_fechados['cohort'] == cohort]
                
                if not cohort_closed_deals.empty and 'Valor do Contrato' in cohort_closed_deals.columns:
                    # Ticket Médio: média do valor dos contratos fechados
                    valores = cohort_closed_deals['Valor do Contrato'].dropna()
                    if len(valores) > 0:
                        cohort_avg_ticket[cohort] = valores.mean()
                        # Ticket Acumulado: soma total dos valores dos contratos fechados
                        cohort_total_ticket[cohort] = valores.sum()
                    else:
                        cohort_avg_ticket[cohort] = 0
                        cohort_total_ticket[cohort] = 0
                else:
                    cohort_avg_ticket[cohort] = 0
                    cohort_total_ticket[cohort] = 0
            
            # Converter para Series e reindexar
            cohort_avg_ticket = pd.Series(cohort_avg_ticket).reindex(all_cohorts, fill_value=0)
            cohort_total_ticket = pd.Series(cohort_total_ticket).reindex(all_cohorts, fill_value=0)
            
            # Agrupar considerando apenas fechados
            df_grouped = df_fechados.groupby(['cohort', 'ano_cohort']).agg(n_customers1=(client_column, 'nunique')).reset_index(drop=False)
            
            # CORREÇÃO: Verificar se há dados agrupados antes de calcular period_number
            if not df_grouped.empty:
                df_grouped['period_number'] = (df_grouped.ano_cohort - df_grouped.cohort).apply(attrgetter('n'))
                # Criar pivot table com os dados de fechados
                cohort_pivot = df_grouped.pivot_table(index='cohort', columns='period_number', values='n_customers1')
            else:
                # Se não houver fechados, criar pivot vazio
                cohort_pivot = pd.DataFrame(index=all_cohorts, columns=[0])
            
            # CORREÇÃO: Reindexar o pivot para incluir todos os cohorts
            cohort_pivot = cohort_pivot.reindex(all_cohorts)
            
            # CORREÇÃO: Preencher NaN com 0 no pivot
            cohort_pivot = cohort_pivot.fillna(0)
            
            # Usar o tamanho total da coorte para calcular a retenção
            cohort_size = cohort_sizes  # Agora já contém todos os cohorts
            
            # CORREÇÃO: Tratar divisão por zero na matriz de retenção
            retention_matrix = cohort_pivot.copy()
            for idx in retention_matrix.index:
                if cohort_size[idx] > 0:
                    retention_matrix.loc[idx] = cohort_pivot.loc[idx] / cohort_size[idx]
                else:
                    retention_matrix.loc[idx] = 0


            # Adicionar toggle para alternar entre porcentagem e valores absolutos
            show_percentage = st.toggle('Mostrar valores em porcentagem', value=False)
            
            # Criar gráfico interativo com Plotly - NOVO: passar métricas adicionais incluindo ticket médio e acumulado
            if show_percentage:
                create_cohort_heatmap_plotly(retention_matrix, cohort_size, cohort_conversion, cohort_closed, cohort_remaining, cohort_lost, cohort_avg_days, cohort_avg_ticket, cohort_total_ticket, is_percentage=True)
            else:
                create_cohort_heatmap_plotly(cohort_pivot, cohort_size, cohort_conversion, cohort_closed, cohort_remaining, cohort_lost, cohort_avg_days, cohort_avg_ticket, cohort_total_ticket, is_percentage=False)

            # NOVO: Adicionar seção de detalhamento de clientes por métrica
            st.markdown("---")
            st.subheader("📋 Detalhamento de Clientes por Métrica")
            
            # Criar seletor de métrica para visualizar detalhes
            metric_to_show = st.selectbox(
                "Selecione a métrica para ver os detalhes dos clientes:",
                options=["Qtd Fechados", "Qtd Perdidos", "Qtd Restante"],
                key="metric_detail_selector"
            )
            
            # Preparar DataFrame baseado na métrica selecionada
            if metric_to_show == "Qtd Fechados":
                status_filter = "Fechado"
                detail_df = df[df['Status'] == status_filter].copy()
                detail_df = detail_df.drop_duplicates(subset=[client_column])
                st.markdown(f"**Total de negócios fechados:** {len(detail_df)}")
                
            elif metric_to_show == "Qtd Perdidos":
                status_filter = "Perdido"
                detail_df = df[df['Status'] == status_filter].copy()
                detail_df = detail_df.drop_duplicates(subset=[client_column])
                st.markdown(f"**Total de negócios perdidos:** {len(detail_df)}")
                
            elif metric_to_show == "Qtd Restante":
                # Clientes que têm "Reunião realizada SQL" mas não têm "Fechado" nem "Perdido"
                reuniao_deals = df[df['Status'] == 'Reunião realizada SQL'][client_column].unique()
                fechado_deals = df[df['Status'] == 'Fechado'][client_column].unique()
                perdido_deals = df[df['Status'] == 'Perdido'][client_column].unique()
                
                # Restante = Reunião - (Fechado + Perdido)
                restante_deals = set(reuniao_deals) - set(fechado_deals) - set(perdido_deals)
                
                detail_df = df[df[client_column].isin(restante_deals)].copy()
                detail_df = detail_df[detail_df['Status'] == 'Reunião realizada SQL']
                detail_df = detail_df.drop_duplicates(subset=[client_column])
                st.markdown(f"**Total de negócios restantes (em aberto):** {len(detail_df)}")
            
            # Exibir DataFrame com informações
            if not detail_df.empty:
                # MODIFICADO: Reorganizar colunas para incluir Programa
                display_cols = [client_column, 'Canal', 'Programa', 'Faturamento Anual', 'Valor do Contrato', 'Status', date_column]
                display_cols = [col for col in display_cols if col in detail_df.columns]
                
                # Formatar valores monetários e datas
                detail_display = detail_df[display_cols].copy()
                
                # Formatar data como DD/MM/YYYY
                if date_column in detail_display.columns:
                    detail_display[date_column] = pd.to_datetime(detail_display[date_column], errors='coerce').dt.strftime('%d/%m/%Y')
                
                if 'Faturamento Anual' in detail_display.columns:
                    detail_display['Faturamento Anual'] = detail_display['Faturamento Anual'].apply(
                        lambda x: f"R$ {x:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.') if pd.notnull(x) and x > 0 else "-"
                    )
                if 'Valor do Contrato' in detail_display.columns:
                    detail_display['Valor do Contrato'] = detail_display['Valor do Contrato'].apply(
                        lambda x: f"R$ {x:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.') if pd.notnull(x) and x > 0 else "-"
                    )
                
                # Adicionar filtros por cohort
                st.markdown("**Filtrar por safra (cohort):**")
                
                # Criar mapeamento de cohorts formatados
                cohort_mapping = {}
                for cohort in sorted(detail_df['cohort'].dropna().unique()):
                    formatted = format_period(cohort, periodo)
                    cohort_mapping[formatted] = cohort
                
                cohort_options = ['Todas'] + list(cohort_mapping.keys())
                selected_cohort_filter = st.selectbox("Selecione a safra:", options=cohort_options, key="cohort_filter")
                
                if selected_cohort_filter != 'Todas':
                    original_cohort = cohort_mapping[selected_cohort_filter]
                    detail_display = detail_display[detail_df['cohort'] == original_cohort]
                
                # Exibir tabela
                st.dataframe(
                    detail_display,
                    use_container_width=True,
                    hide_index=True
                )
                
                # Botão para download
                csv = detail_display.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label=f"⬇️ Baixar lista de {metric_to_show} em CSV",
                    data=csv,
                    file_name=f"detalhamento_{metric_to_show.lower().replace(' ', '_')}.csv",
                    mime="text/csv"
                )
            else:
                st.info("Nenhum cliente encontrado para esta métrica com os filtros aplicados.")


        except Exception as e:
            st.error(f"Erro na análise: {str(e)}")
            import traceback
            st.error(traceback.format_exc())
    else:
        st.info("Aguardando o upload do Excel do Hubspot.")


# MODIFICAÇÃO 3: Função create_cohort_heatmap_plotly() - adicionar os novos parâmetros e subplots
def create_cohort_heatmap_plotly(matrix, cohort_size, cohort_conversion, cohort_closed, cohort_remaining, cohort_lost, cohort_avg_days, cohort_avg_ticket, cohort_total_ticket, is_percentage=True):
    # CORREÇÃO: Obter o tipo de período do session_state ou inferir da matriz
    # Vamos inferir o tipo de período a partir do formato dos índices
    if len(matrix.index) > 0:
        first_cohort = str(matrix.index[0])
        if 'Q' in first_cohort:
            periodo_type = 'Q'
        elif len(first_cohort) == 4:  # Apenas ano (ex: "2025")
            periodo_type = 'Y'
        else:
            periodo_type = 'M'
    else:
        periodo_type = 'M'  # Default
    
    # Preparar dados para o gráfico - FORMATAR OS LABELS DAS SAFRAS
    cohort_labels = [format_period(cohort, periodo_type) for cohort in matrix.index]
    
    period_cols = list(matrix.columns)
    
    try:
        period_ints = [int(x) for x in period_cols]
    except Exception:
        period_ints = list(range(len(period_cols)))

    if period_ints:
        min_p = min(period_ints)
        max_p = max(period_ints)
    else:
        min_p = 0
        max_p = 0
    desired_max = max(max_p, 5)
    full_periods = list(range(min_p, desired_max + 1))

    matrix = matrix.reindex(columns=full_periods)

    period_cols = full_periods
    period_labels = [str(col) for col in period_cols]
    
    sales_cycle_labels = []
    for period in period_cols:
        if period == 0:
            sales_cycle_labels.append("<30")
        elif period == 1:
            sales_cycle_labels.append("30-60")
        else:
            inicio = 30 * period
            fim = 30 * (period + 1)
            sales_cycle_labels.append(f"{inicio}-{fim}")

    selected_metrics = st.multiselect(
        "Selecione as métricas adicionais para exibir:", 
        options=["Tx de Conversão", "Qtd Fechados", "Qtd Perdidos", "Qtd Restante", "Tempo Médio", "Ticket Médio", "Ticket Acumulado"],
        default=["Tx de Conversão", "Qtd Fechados", "Qtd Perdidos", "Qtd Restante", "Tempo Médio", "Ticket Médio", "Ticket Acumulado"],
        key="additional_metrics"
    )

    num_metrics = len(selected_metrics)
    total_cols = 2 + num_metrics

    fixed_widths = [0.05, 0.65]
    metric_widths = [0.06] * num_metrics
    column_widths = fixed_widths + metric_widths

    fig = make_subplots(
        rows=1, 
        cols=total_cols,
        column_widths=column_widths,
        shared_yaxes=True,
        horizontal_spacing=0.005,
    )
    
    # Coluna 1: Tamanho da safra
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
    
    # Coluna 2: Matriz principal - MODIFICADO PARA TRATAR ZEROS COMO TRANSPARENTES
    # Criar textos para exibição e substituir valores 0 por NaN apenas para visualização
    text_matrix = []
    hover_matrix = []
    
    # NOVO: Criar cópia da matriz e substituir 0 por NaN APENAS para visualização
    z_values = matrix.copy()
    z_values = z_values.replace(0, np.nan)
    
    for i in range(matrix.shape[0]):
        text_row = []
        hover_row = []
        for j in range(matrix.shape[1]):
            # Usar a matriz ORIGINAL (com zeros) para gerar texto e hover
            if not pd.isna(matrix.iloc[i, j]) and matrix.iloc[i, j] > 0:
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
    
    # Calcular zmax seguro
    max_val = matrix.max().max()
    if pd.isna(max_val) or max_val == 0:
        max_val = 1  # Default para evitar erro
    
    fig.add_trace(
        go.Heatmap(
            z=z_values,  # Usar matriz com NaN para cores
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
            zmax=1 if is_percentage else max_val,
            connectgaps=False,
            hoverongaps=False,
            showscale=True,
            zauto=False,
            zmid=0.5,
        ),
        row=1, col=2
    )
    
    base_col = 3

    # Tx de Conversão
    if "Tx de Conversão" in selected_metrics:
        conversion_values = cohort_conversion.values.reshape(-1, 1)
        conversion_text = [[f"{val:.1f}%" for val in conversion_values.flatten()]]
        conversion_text = list(zip(*conversion_text))
        
        col_position = base_col + selected_metrics.index("Tx de Conversão")
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
            row=1, col=col_position
        )
    
    # Qtd Fechados
    if "Qtd Fechados" in selected_metrics:
        closed_values = cohort_closed.values.reshape(-1, 1)
        closed_text = [[str(int(val)) for val in closed_values.flatten()]]
        closed_text = list(zip(*closed_text))

        col_position = base_col + selected_metrics.index("Qtd Fechados")
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
            row=1, col=col_position
        )
    
    # Qtd Perdidos
    if "Qtd Perdidos" in selected_metrics:
        lost_values = cohort_lost.values.reshape(-1, 1)
        lost_text = [[str(int(val)) for val in lost_values.flatten()]]
        lost_text = list(zip(*lost_text))

        col_position = base_col + selected_metrics.index("Qtd Perdidos")
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
                hovertext=[[f'Perdidos: {int(val)}' for val in lost_values.flatten()]],
            ),
            row=1, col=col_position
        )

    # Qtd Restante
    if "Qtd Restante" in selected_metrics:
        remaining_values = cohort_remaining.values.reshape(-1, 1)
        remaining_text = [[str(int(val)) for val in remaining_values.flatten()]]
        remaining_text = list(zip(*remaining_text))

        col_position = base_col + selected_metrics.index("Qtd Restante")
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
            row=1, col=col_position
        )

    # Tempo Médio
    if "Tempo Médio" in selected_metrics:
        avg_days_values = cohort_avg_days.values.reshape(-1, 1)
        avg_days_text = [[f"{val:.0f}d" if val > 0 else "-" for val in avg_days_values.flatten()]]
        avg_days_text = list(zip(*avg_days_text))

        col_position = base_col + selected_metrics.index("Tempo Médio")
        fig.add_trace(
            go.Heatmap(
                z=avg_days_values,
                x=['Dias'],
                y=cohort_labels,
                text=avg_days_text,
                texttemplate="%{text}",
                textfont={"size": 16},
                colorscale=[[0, 'white'], [1, 'white']],
                showscale=False,
                hoverinfo='text',
                hovertext=[[f'<b>Tempo médio: {val:.0f} dias</b>' if val > 0 else '<b>Sem dados</b>' for val in avg_days_values.flatten()]],
            ),
            row=1, col=col_position
        )

    # NOVO: Ticket Médio
    if "Ticket Médio" in selected_metrics:
        avg_ticket_values = cohort_avg_ticket.values.reshape(-1, 1)
        # Formatar em milhares (K)
        avg_ticket_text = [[f"R$ {val/1000:.0f}K" if val > 0 else "-" for val in avg_ticket_values.flatten()]]
        avg_ticket_text = list(zip(*avg_ticket_text))

        col_position = base_col + selected_metrics.index("Ticket Médio")
        fig.add_trace(
            go.Heatmap(
                z=avg_ticket_values,
                x=['Tk Médio'],
                y=cohort_labels,
                text=avg_ticket_text,
                texttemplate="%{text}",
                textfont={"size": 16},
                colorscale=[[0, 'white'], [1, 'white']],
                showscale=False,
                hoverinfo='text',
                hovertext=[[f'<b>Ticket Médio: R$ {val:,.2f}</b>'.replace(',', '_').replace('.', ',').replace('_', '.') if val > 0 else '<b>Sem dados</b>' for val in avg_ticket_values.flatten()]],
            ),
            row=1, col=col_position
        )

    # NOVO: Ticket Acumulado
    if "Ticket Acumulado" in selected_metrics:
        total_ticket_values = cohort_total_ticket.values.reshape(-1, 1)
        # Formatar em milhares (K)
        total_ticket_text = [[f"R$ {val/1000:.0f}K" if val > 0 else "-" for val in total_ticket_values.flatten()]]
        total_ticket_text = list(zip(*total_ticket_text))

        col_position = base_col + selected_metrics.index("Ticket Acumulado")
        fig.add_trace(
            go.Heatmap(
                z=total_ticket_values,
                x=['Tk Acum.'],
                y=cohort_labels,
                text=total_ticket_text,
                texttemplate="%{text}",
                textfont={"size": 16},
                colorscale=[[0, 'white'], [1, 'white']],
                showscale=False,
                hoverinfo='text',
                hovertext=[[f'<b>Ticket Acumulado: R$ {val:,.2f}</b>'.replace(',', '_').replace('.', ',').replace('_', '.') if val > 0 else '<b>Sem dados</b>' for val in total_ticket_values.flatten()]],
            ),
            row=1, col=col_position
        )

    fig.update_layout(
        height=700,
        width=1400,
        yaxis=dict(
            title='',
            autorange='reversed',
            type='category',  # NOVO: Forçar tipo categoria para evitar conversão automática de datas
        ),
        xaxis2=dict(
            title='# períodos',
            tickmode='array',
            tickvals=list(range(len(period_labels))),
            ticktext=period_labels,
            side='top',
        ),
    )
    
    fig.update_xaxes(
        title='# sales cycle (dias)',
        tickmode='array',
        tickvals=list(range(len(sales_cycle_labels))),
        ticktext=sales_cycle_labels,
        side='bottom',
        row=1, col=2
    )

    fig.update_layout(margin=dict(l=50, r=50, t=100, b=100))
    fig.update_xaxes(showgrid=False)
    fig.update_xaxes(showgrid=False)
    
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
