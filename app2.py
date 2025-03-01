from flask import Flask, render_template, request, jsonify, flash, redirect, url_for, session
import pandas as pd
import pyodbc
import json
import warnings
import plotly.graph_objs as go
import plotly
from datetime import datetime, timedelta
import calendar
import os
from rich import print as rprint
import click
import logging
import re
import difflib
from bs4 import BeautifulSoup

import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords

def verificar_e_instalar_nltk():
    """Verifica se os pacotes do NLTK estão instalados e os baixa apenas se necessário."""
    required_packages = {
        "punkt": "tokenizers/punkt",
        "punkt_tab": "tokenizers/punkt_tab",
        "stopwords": "corpora/stopwords",
        "averaged_perceptron_tagger": "taggers/averaged_perceptron_tagger",
        "averaged_perceptron_tagger_eng": "taggers/averaged_perceptron_tagger_eng"
    }

    for package, path in required_packages.items():
        try:
            nltk.data.find(path)
            # print(f"✔ {package} já está instalado.")
        except LookupError:
            # print(f"⬇ Baixando {package}...")
            nltk.download(package)


log = logging.getLogger('werkzeug')
log.setLevel(logging.ERROR)

def secho(text, file=None, nl=None, err=None, color=None, **styles):
    pass

def echo(text, file=None, nl=None, err=None, color=None, **styles):
    pass

click.echo = echo
click.secho = secho

warnings.filterwarnings('ignore')

app = Flask(__name__) 
app.secret_key = "testeunique"

# Configuração da conexão com o banco de dados Access 
def get_db_connection():
    try:
        conn_str = (
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            r"DBQ=C:\\Users\\Henrique\\Downloads\\Controle.accdb"
        )
        conn = pyodbc.connect(conn_str)
        return conn
    except pyodbc.Error as e:
        print(f"Erro ao conectar ao banco de dados: {e}")
        return None

# Dicionário para os meses em português
meses_dict = {
    "Janeiro": "01", "Fevereiro": "02", "Março": "03", "Abril": "04",
    "Maio": "05", "Junho": "06", "Julho": "07", "Agosto": "08",
    "Setembro": "09", "Outubro": "10", "Novembro": "11", "Dezembro": "12"
}

# Dicionário de cores e marcadores para cada tipo de presença
color_marker_map = {
    'OK': {'cor': '#494949', 'marker': 'circle'},
    'FALTA': {'cor': '#FF5733', 'marker': 'x'},
    'ATESTADO': {'cor': '#FFC300', 'marker': 'diamond'},
    'FOLGA': {'cor': '#233F7B', 'marker': 'diamond'},
    'CURSO': {'cor': '#a12a8f', 'marker': 'star'},
    'FÉRIAS': {'cor': '#a5a5a5', 'marker': 'square'},
    'ALPHAVILLE':{'cor': '#76A9B7', 'marker': 'square'},
    'LICENÇA':{'cor': '#632aa1', 'marker': 'diamond'},
}

# Dicionário de saudações válidas para o chatbot
saudacoes_validas = ["olá", 
                     "oi", "e aí", "opa", "fala", "alô", "olá chat", "bom dia chat", "boa noite chat", "boa tarde chat", "salve", "olá tudo bem?", "oi tudo bem?", "e aí tudo bem?", "opa tudo bem?", "fala tudo bem?", "alô tudo bem?", "olá chat tudo bem?", "bom dia chat tudo bem?", "boa noite chat tudo bem?", "boa tarde chat tudo bem?", "salve tudo bem?"]

# Perguntas e respostas para o chatbot
perguntas_respostas = {
    "adicionar": {
        "presenca": {
            "perguntas": [
                "como adiciono uma presença?", "quero registrar uma presença", "como inserir uma presença?",
                "onde adiciono presença?", "como faço para cadastrar presença?", "como faço para marcar presença?",
                "adicionar presença", "inserir presença", "onde faço para adicionar presença?", "como marcar uma presença?",
                "onde posso registrar uma presença?", "como coloco uma presença?", "como registrar presença?",
                "quero adicionar um nome na presença", "como faço para incluir uma presença?",
                "como faço para adicionar presença no sistema?", "como faço para salvar uma presença?",
                "quero registrar um funcionário presente", "onde posso lançar presença?"
            ],
            "resposta": "Para adicionar presença, acesse a página 'Adicionar Presença', selecione o nome, data e tipo de presença e clique em 'Salvar'."
        },
        "nome": {
            "perguntas": [
                "como adiciono um nome?", "quero registrar um novo nome", "como inserir um nome?",
                "onde adiciono nome?", "como faço para cadastrar um novo nome?", "como faço para incluir um nome?",
                "como insiro um novo nome no sistema?", "como posso cadastrar um nome?", "onde adiciono uma nova pessoa?",
                "como coloco um nome no sistema?", "quero incluir um nome no cadastro", "onde faço o registro de nome?",
                "onde posso adicionar um novo colaborador?", "como cadastrar novo usuário?"
            ],
            "resposta": "Para adicionar um nome, vá até a página 'Adicionar Presença', digite o nome e clique em 'Salvar'."
        },
        "empresa": {
            "perguntas": [
                "como adiciono uma empresa?", "quero registrar uma nova empresa ao site", "como inserir uma empresa?",
                "onde adiciono uma empresa?", "como faço para cadastrar uma empresa?", "como faço para adicionar uma empresa?",
                "onde posso registrar uma empresa?", "adicionar empresa", "inserir empresa",
                "onde cadastro uma nova empresa?", "como faço para incluir uma empresa no sistema?",
                "quero adicionar uma nova organização", "quero incluir uma nova empresa", "como cadastrar empresa?",
                "onde adiciono um novo CNPJ?", "como faço para cadastrar uma nova firma?"
            ],
            "resposta": "Para adicionar uma empresa, acesse a página 'Adicionar Presença', digite o nome da empresa e clique em 'Salvar'."
        }
    },
    "remover": {
        "presenca": {
            "perguntas": [
                "como remover uma presença?", "quero excluir uma presença", "como apagar uma presença?",
                "onde posso deletar uma presença?", "remover presença", "excluir presença", "apagar presença",
                "deletar presença", "como cancelo uma presença?", "como retiro uma presença?",
                "como desfazer um lançamento de presença?", "como faço para corrigir um erro na presença?",
                "remover presença de um funcionário", "quero cancelar uma presença já registrada"
            ],
            "resposta": "Para remover presença, acesse a página 'Controle de Presença', selecione a data e clique em 'Remover'."
        },
        "nome": {
            "perguntas": [
                "como remover um nome?", "quero excluir um nome do controle", "como apagar um nome?",
                "onde posso deletar um nome?", "remover nome", "excluir nome", "apagar nome", "deletar nome",
                "como cancelo um nome?", "quero excluir um colaborador", "como faço para retirar um nome?",
                "remover funcionário do sistema", "como eliminar um nome cadastrado?", "onde posso excluir um usuário?"
            ],
            "resposta": "Para remover um nome, acesse a página 'Controle de Presença', selecione o nome e clique em 'Remover'."
        },
        "empresa": {
            "perguntas": [
                "como remover uma empresa?", "quero excluir uma empresa do controle", "como apagar uma empresa?",
                "onde posso deletar uma empresa?", "remover empresa", "excluir empresa", "apagar empresa",
                "deletar empresa", "como cancelo uma empresa?", "como faço para remover um CNPJ?",
                "onde retiro uma empresa cadastrada?", "como faço para desativar uma empresa?",
                "quero excluir uma firma do sistema", "onde faço a remoção de uma empresa cadastrada?"
            ],
            "resposta": "Para remover uma empresa, acesse a página 'Controle de Presença', selecione a empresa e clique em 'Remover'."
        }
    },
    "filtrar": {
        "perguntas": [
            "como filtrar presenças?", "quero buscar um nome específico", "como faço para ver as presenças de um mês?",
            "como aplico um filtro nas presenças?", "filtrar presença", "quero pesquisar uma presença",
            "como vejo quem esteve presente?", "quero encontrar um nome", "como posso filtrar os registros?",
            "onde aplico um filtro para ver presenças?", "existe um jeito de filtrar as presenças?",
            "como faço para listar presenças de um período?", "onde vejo registros por data?",
            "como encontro um funcionário pelo nome?", "como ver lista de presenças de um mês específico?",
            "como filtrar funcionários por empresa?", "onde posso ver um histórico de presenças?"
        ],
        "resposta": "Para filtrar, utilize os campos de nome, mês, tipo de presença e ano na página principal."
    }
}

# Dicionário para mapear meses abreviados, completos e numéricos corretamente
meses_map = {
    "jan": "01", "janeiro": "01",
    "fev": "02", "fevereiro": "02",
    "mar": "03", "março": "03",
    "abr": "04", "abril": "04",
    "mai": "05", "maio": "05",
    "jun": "06", "junho": "06",
    "jul": "07", "julho": "07",
    "ago": "08", "agosto": "08",
    "set": "09", "setembro": "09",
    "out": "10", "outubro": "10",
    "nov": "11", "novembro": "11",
    "dez": "12", "dezembro": "12"
}

# Criar uma lista de nomes completos de meses para verificar
meses_completos = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", 
    "agosto", "setembro", "outubro", "novembro", "dezembro"
]

# Dicionário para converter tipos de frequência do plural para o singular
frequencia_plural_para_singular = {
    "oks": "ok", "faltas": "falta", "atestados": "atestado",
    "cursos": "curso", "folgas": "folga", "ferias": "férias",
    "férias": "férias", "licenças": "licença",
}

LISTAGEM_NOMES = [
    "nomes disponivel", "lista de nomes", "quais são os nomes cadastrados",
    "quero ver a listagem de nomes", "mostrar todos os nomes", "exibir nomes",
    "quais são os funcionários cadastrados", "me mostre os nomes disponíveis",
    "listar nomes", "quem está cadastrado no sistema", "mostrar nomes", 
    "exibir listagem de nomes", "onde vejo a lista de nomes", "quero ver os nomes",
    "quais são os nomes registrados"
]

@app.route("/capturar_largura_tela", methods=["POST"])
def capturar_largura_tela():
    screen_width = request.json.get("screenWidth")
    screen_width = int(screen_width)
    
    # Defina a largura do gráfico com base na largura da tela recebida
    if screen_width <= 1064:
        largura_grafico = 800
    elif 1065 <= screen_width <= 1600:
        largura_grafico = 450
    elif screen_width >= 1920 and screen_width <= 1930:
        largura_grafico = 500
    elif screen_width >= 1930:
        largura_grafico = 1000
    else:
        largura_grafico = 500
    
    # Armazena a largura do gráfico na sessão
    session['larguraGrafico'] = largura_grafico
    
    return jsonify({"larguraGrafico": largura_grafico})

@app.route("/", methods=["GET", "POST"])
def index():
    conn = get_db_connection()
    if not conn:
        flash("Erro ao conectar ao banco de dados.", "error")
        return redirect(url_for('index'))
    
    # Consultar sites
    query_sites = "SELECT DISTINCT Sites FROM Site"
    sites = pd.read_sql(query_sites, conn)['Sites'].tolist()

    # Captura os valores dos filtros
    selected_site = request.form.get("site") or session.get('selected_site')
    selected_empresa = request.form.get("empresa") or session.get('selected_empresa')
    selected_ano = request.form.getlist("ano")  # Captura o valor do ano selecionado
    largura_grafico = session.get('larguraGrafico') 
    # Salva os valores na sessão
    if selected_site:
        session['selected_site'] = selected_site
    if selected_empresa:
        session['selected_empresa'] = selected_empresa
    
    # Captura o intervalo de datas do formulário em vez de request.json
    date_range = request.form.get("dateRange")
    # Processa o intervalo de datas
    if date_range:
        try:
            start_date_str, end_date_str = date_range.split(" to ")
            
            start_date = pd.to_datetime(start_date_str, dayfirst=True, errors='coerce')
            end_date = pd.to_datetime(end_date_str, dayfirst=True, errors='coerce')
            # print(f"Intervalo de datas: {start_date} \t {end_date}")
        except Exception as e:
            print(f"Erro ao processar o intervalo de datas: {e}")
            start_date = None
            end_date = None
    else:
        start_date = None
        end_date = None


    selected_nomes = request.form.getlist("nomes")
    selected_meses = request.form.getlist("meses")
    selected_presenca = request.form.getlist("presenca")

    # Buscar os anos distintos no banco de dados (Ajuste para Microsoft Access)
    query_anos = """
        SELECT Ano 
        FROM (SELECT DISTINCT YEAR(Data) as Ano FROM Controle) AS subquery 
        ORDER BY Ano DESC
    """
    anos = pd.read_sql(query_anos, conn)['Ano'].astype(str).tolist()


    empresas = []
    if selected_site:
        empresas = get_empresas(get_site_id(selected_site))
        site_id = get_site_id(selected_site)
        

    # Inicializa a tabela como vazia
    df = pd.DataFrame(columns=['Nome', 'Presenca', 'Data'])
    pres = pd.read_sql("SELECT DISTINCT Presenca FROM Presenca", conn)['Presenca'].tolist()
    conn.close()
    # Variáveis para os gráficos (inicializando com None)
    pie_chart_data = None
    scatter_chart_data = None
    stacked_bar_chart_data = None
    total_dias_registrados = 0
    total_ok = 0
    total_faltas = 0
    total_atestados = 0
    nomes = []
    
    # Executa a consulta SQL somente se site e empresa forem selecionados
    if selected_site and selected_empresa:
        try:
            conn = get_db_connection()
            empresa_id = get_empresa_id(selected_empresa, empresas)
            siteempresa_id = get_siteempresa_id(site_id, empresa_id)
            nomes = get_nomes(siteempresa_id, ativos=True)
            query = """
            SELECT Nome.Nome, Presenca.Presenca, Controle.Data
            FROM (((Controle
            INNER JOIN Nome ON Controle.id_Nome = Nome.id_Nomes)
            INNER JOIN Presenca ON Controle.id_Presenca = Presenca.id_Presenca)
            INNER JOIN Site_Empresa ON Controle.id_SiteEmpresa = Site_Empresa.id_SiteEmpresa)
            WHERE Site_Empresa.id_Sites = ? AND Site_Empresa.id_Empresas = ?
            """
            query_params = [get_site_id(selected_site), get_empresa_id(selected_empresa, empresas)]
                
            # Verifica se algum ano foi selecionado
            if selected_ano and len(selected_ano) > 0:
                if len(selected_ano) == 1:
                    # Apenas um ano selecionado -> Filtro direto
                    query += " AND YEAR(Controle.Data) = ?"
                    query_params.append(selected_ano[0])  # Adiciona o único ano
                else:
                    # Vários anos selecionados -> Criando placeholders dinâmicos
                    anos_placeholder = ",".join(["?"] * len(selected_ano))
                    query += f" AND YEAR(Controle.Data) IN ({anos_placeholder})"
                    query_params.extend(selected_ano)  # Adiciona todos os anos na lista de parâmetros



            cursor = conn.cursor()
            cursor.execute(query, query_params)
            rows = cursor.fetchall()

            # Verificar se há dados retornados
            if rows:
                df = pd.DataFrame([list(row) for row in rows], columns=['Nome', 'Presenca', 'Data'])

                # Converte a coluna Data para datetime
                df['Data'] = pd.to_datetime(df['Data'], format='%Y-%m-%d %H:%M:%S')
                if start_date and end_date:
                    df = df[(df['Data'] >= start_date) & (df['Data'] <= end_date)]

                # Aplicar filtros adicionais
                if selected_nomes:
                    df = df[df['Nome'].isin(selected_nomes)]
                if selected_presenca:
                    df = df[df['Presenca'].isin(selected_presenca)]
                if selected_meses:
                    selected_meses_numeric = [meses_dict[mes] for mes in selected_meses]
                    df = df[df['Data'].dt.strftime('%m').isin(selected_meses_numeric)]
                if selected_ano and len(selected_ano) > 0:
                    df = df[df['Data'].dt.year.astype(str).isin(selected_ano)]

                # Gera uma lista contínua de datas entre o menor e o maior valor de data
                min_data = df['Data'].min()
                max_data = df['Data'].max()
                datas_continuas = pd.date_range(min_data, max_data).to_list()

                # Cria uma nova DataFrame com todas as combinações possíveis de nomes e datas contínuas
                nomes_unicos = df['Nome'].unique()
                df_continuo = pd.MultiIndex.from_product([nomes_unicos, datas_continuas], names=['Nome', 'Data']).to_frame(index=False)

                # Converte ambas as colunas 'Data' para datetime para garantir a compatibilidade no merge
                df_continuo['Data'] = pd.to_datetime(df_continuo['Data'])
                df['Data'] = pd.to_datetime(df['Data'])

                # Faz o merge do DataFrame original com o DataFrame contínuo
                df_merge = pd.merge(df_continuo, df, on=['Nome', 'Data'], how='left')

                # 🔹 Aplicando filtro explícito de ano no DataFrame dos gráficos
                if selected_ano and len(selected_ano) > 0:
                    df_merge = df_merge[df_merge['Data'].dt.year.astype(str).isin(selected_ano)]

                # Preenche valores ausentes com "invisível" ou algum valor placeholder
                df_merge['Presenca'] = df_merge['Presenca'].fillna('invisível')

                # Gráfico de dispersão
                fig_dispersao = go.Figure()

                for presenca, info in color_marker_map.items():
                    df_tipo = df_merge[df_merge['Presenca'].str.upper() == presenca]
                    if not df_tipo.empty:
                        fig_dispersao.add_trace(go.Scatter(
                            x=df_tipo['Data'],
                            y=df_tipo['Nome'],
                            mode='markers',
                            marker=dict(color=info['cor'], symbol=info['marker'], size=10),
                            name=presenca
                        ))

                # Adicionar os pontos invisíveis para garantir o espaçamento correto
                df_invisivel = df_merge[df_merge['Presenca'] == 'invisível']
                fig_dispersao.add_trace(go.Scatter(
                    x=df_invisivel['Data'],
                    y=df_invisivel['Nome'],
                    mode='markers',
                    marker=dict(color='rgba(0,0,0,0)', size=10),  # Invisível
                    name='invisível',
                    showlegend=False  # Não mostrar na legenda
                ))

                if selected_meses:
                    # Customizando o layout do gráfico de dispersão
                    fig_dispersao.update_layout(
                        title={
                            'text': "Gráfico de Dispersão de Presenças",
                            'x': 0.5,
                            'xanchor': 'center',
                            'yanchor': 'top',
                            'font': {'size': 24}
                        },
                        xaxis=dict(
                            showgrid=False,
                            gridcolor='lightgray',
                            tickformat='%d/%m/%Y',  # Formata as datas no eixo X como dd/mm/yyyy
                            dtick=86400000
                        ),
                        yaxis=dict(showgrid=False, gridcolor='lightgray'),
                        font=dict(color='#000000'),
                        plot_bgcolor='rgba(0,0,0,0)',
                        paper_bgcolor='rgba(0,0,0,0)',
                        hovermode='closest'
                    )
                else:
                      # Customizando o layout do gráfico de dispersão
                    fig_dispersao.update_layout(
                        title={
                            'text': "Gráfico de Dispersão de Presenças",
                            'x': 0.5,
                            'xanchor': 'center',
                            'yanchor': 'top',
                            'font': {'size': 24}
                        },
                        xaxis=dict(
                            showgrid=False,
                            gridcolor='lightgray',
                            tickformat='%d/%m/%Y',  # Formata as datas no eixo X como dd/mm/yyyy
                            # dtick=86400000
                        ),
                        yaxis=dict(showgrid=False, gridcolor='lightgray'),
                        font=dict(color='#000000'),
                        plot_bgcolor='rgba(0,0,0,0)',
                        paper_bgcolor='rgba(0,0,0,0)',
                        hovermode='closest'
                    )

                # Converte o gráfico de dispersão para JSON para renderizar no HTML
                scatter_chart_data = json.dumps(fig_dispersao, cls=plotly.utils.PlotlyJSONEncoder)

                # Gráfico de Pizza (usando Plotly)
                df_presenca = df.groupby('Presenca').size().reset_index(name='counts')
                labels = df_presenca['Presenca'].str.upper().tolist()  # Tipos de presença em maiúsculas
                values = df_presenca['counts'].tolist()    # Contagens de cada presença

                # Mapeamento das cores para o gráfico de pizza
                colors = [color_marker_map[label]['cor'] if label in color_marker_map else '#999999' for label in labels]

                # Criação do gráfico de pizza com Plotly
                fig_pie = go.Figure(data=[go.Pie(labels=labels, values=values, textinfo='label+percent', hole=0.3, marker=dict(colors=colors))])

                # Definir layout do gráfico de pizza
                fig_pie.update_layout(
                    title={
                        'text': "Distribuição de Presença",
                        'x': 0.5,  # Centraliza o título
                        'xanchor': 'center',
                        'yanchor': 'top',
                        'font': {'size': 24}  # Altera o tamanho da fonte do título
                    },
                    showlegend=True,
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)'
                )

                # Converte o gráfico de pizza para JSON
                pie_chart_data = json.dumps(fig_pie, cls=plotly.utils.PlotlyJSONEncoder)

                # Gráfico de Barras Empilhadas
                df['Presenca'] = df['Presenca'].str.upper()
                df_agrupado = df.groupby(['Nome', 'Presenca']).size().reset_index(name='counts')
                barras = []

                for presenca in df_agrupado['Presenca'].unique():
                    df_presenca = df_agrupado[df_agrupado['Presenca'] == presenca]
                    barra = go.Bar(
                        x=df_presenca['Nome'],
                        y=df_presenca['counts'],
                        name=presenca,
                        marker=dict(color=color_marker_map[presenca]['cor']),
                        text=df_presenca['counts'],
                        textposition='inside'
                    )
                    barras.append(barra)

                layout = go.Layout(
                    title={
                        'text': "Nomes x Presença",
                        'x': 0.5,  # Centraliza o título
                        'xanchor': 'center',
                        'yanchor': 'top',
                        'font': {'size': 24}  # Altera o tamanho da fonte do título
                    },
                    barmode='stack',
                    width=largura_grafico,
                    xaxis=dict(title='Nome', showgrid=False),
                    yaxis=dict(title='Contagem de Presença', showgrid=False),
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    font=dict(color='#000000')
                )

                fig_barras_empilhadas = go.Figure(data=barras, layout=layout)
                stacked_bar_chart_data = json.dumps(fig_barras_empilhadas, cls=plotly.utils.PlotlyJSONEncoder)

                # Contagem de dias únicos para o resumo
                total_dias_registrados = df['Data'].nunique()  # Contagem de dias únicos
                total_ok = df[df['Presenca'].str.upper() == 'OK'].shape[0]  # Contagem de OK
                total_faltas = df[df['Presenca'].str.upper() == 'FALTA'].shape[0]  # Contagem de FALTAS
                total_atestados = df[df['Presenca'].str.upper() == 'ATESTADO'].shape[0]  # Contagem de ATESTADOS

                # Formatar a coluna 'Data' para o formato 'dd/mm/yyyy' para a tabela
                df['Data'] = df['Data'].dt.strftime('%d/%m/%Y')
        
            nomes = get_nomes(siteempresa_id, ativos=True) + get_nomes(siteempresa_id, ativos=False) if siteempresa_id else []

        except Exception as e:
            print(f"Erro ao consultar ou criar DataFrame: {e}")
        conn.close()
    return render_template(
        "index.html",
        sites=sites,
        empresas=[e[1] for e in empresas],
        nomes = nomes,
        meses=meses_dict.keys(),
        presencas=pres,
        selected_site=selected_site,
        selected_empresa=selected_empresa,
        selected_nomes=selected_nomes,
        selected_meses=selected_meses,
        selected_presenca=selected_presenca,
        data=df,  # Agora com as datas formatadas para dd/mm/yyyy
        pie_chart_data=pie_chart_data,
        scatter_chart_data=scatter_chart_data,  # Gráfico de dispersão com datas formatadas
        stacked_bar_chart_data=stacked_bar_chart_data,
        total_dias_registrados=total_dias_registrados,
        total_ok=total_ok,
        total_faltas=total_faltas,
        total_atestados=total_atestados,
        color_marker_map=color_marker_map,
        anos=anos,
        selected_ano=selected_ano,
    )

def get_site_id(site_name):
    conn = get_db_connection()
    if not conn:
        return None
    cursor = conn.cursor()
    cursor.execute("SELECT id_Site FROM Site WHERE Sites = ?", (site_name,))
    result = cursor.fetchone()
    conn.close()
    return result[0] if result else None


def get_empresas(site_id):
    conn = get_db_connection()
    if not conn:
        return []
    cursor = conn.cursor()
    query = """
    SELECT Empresa.id_Empresa, Empresa.Empresas
    FROM Site_Empresa
    INNER JOIN Empresa ON Site_Empresa.id_Empresas = Empresa.id_Empresa
    WHERE Site_Empresa.id_Sites = ? AND Site_Empresa.Ativo = True
    """
    cursor.execute(query, (site_id,))
    empresas = [(row[0], row[1]) for row in cursor.fetchall()]
    conn.close()
    return empresas


def get_empresas_inativas(site_id):
    conn = get_db_connection()
    if not conn:
        return []
    cursor = conn.cursor()
    query = """
    SELECT Empresa.id_Empresa, Empresa.Empresas
    FROM Site_Empresa
    INNER JOIN Empresa ON Site_Empresa.id_Empresas = Empresa.id_Empresa
    WHERE Site_Empresa.id_Sites = ? AND Site_Empresa.Ativo = False
    """
    cursor.execute(query, (site_id,))
    empresas = [(row[0], row[1]) for row in cursor.fetchall()]
    conn.close()
    return empresas


def get_empresa_id(empresa_nome, empresas):
    for empresa in empresas:
        if empresa[1] == empresa_nome:
            return empresa[0]
    return None


def get_siteempresa_id(site_id, empresa_id):
    conn = get_db_connection()
    if not conn:
        return None
    cursor = conn.cursor()
    query = "SELECT id_SiteEmpresa FROM Site_Empresa WHERE id_Sites = ? AND id_Empresas = ? AND Ativo = True"
    cursor.execute(query, (site_id, empresa_id))
    result = cursor.fetchone()
    conn.close()
    return result[0] if result else None

def get_nomes(siteempresa_id, ativos=True):
    """Obtém os nomes associados ao ID_SiteEmpresas, filtrando por ativos se solicitado."""
    conn = get_db_connection()
    if not conn:
        return []
    cursor = conn.cursor()
    query = "SELECT Nome.Nome FROM Nome WHERE id_SiteEmpresa = ?"

    if ativos:
        query += " AND Ativo = True"
    else:
        query += " AND Ativo = False"

    cursor.execute(query, (siteempresa_id,))
    nomes = [row[0] for row in cursor.fetchall()]
    conn.close()
    return nomes

# Tabela de presença da pagina de (adicionar presença)
def fetch_registros_mes(site_id, empresa_id, current_month, current_year):
    """
    Busca os registros do mês e ano atuais no banco de dados Access.
    """
    try:
        conn = get_db_connection()
        if not conn:
            raise ConnectionError("Erro ao conectar ao banco de dados.")

        # Consulta para pegar os registros do mês e ano atual
        query = """
            SELECT Nome.Nome, Presenca.Presenca, Controle.Data
            FROM (((Controle
            INNER JOIN Nome ON Controle.id_Nome = Nome.id_Nomes)
            INNER JOIN Presenca ON Controle.id_Presenca = Presenca.id_Presenca)
            INNER JOIN Site_Empresa ON Controle.id_SiteEmpresa = Site_Empresa.id_SiteEmpresa)
            WHERE Site_Empresa.id_Sites = ? AND Site_Empresa.id_Empresas = ?
            AND MONTH(Controle.Data) = ? AND YEAR(Controle.Data) = ?
        """
        cursor = conn.cursor()
        cursor.execute(query, (site_id, empresa_id, current_month, current_year))
        registros_mes_atual = cursor.fetchall()  # Pega os registros

        # Formatar como lista de tuplas (Nome, Presença, Data)
        registros_formatados = [
            (row[0], row[1], row[2]) for row in registros_mes_atual
        ]
        conn.close()
        return registros_formatados

    except Exception as e:
        print(f"Erro ao buscar registros do mês: {e}")
        return []

@app.route('/adicionar-presenca', methods=['GET', 'POST'])
def adiciona_presenca():
    conn = get_db_connection()
    if not conn:
        flash("Erro ao conectar ao banco de dados.", "error")
        return redirect(url_for('index'))

    # Consultar sites e presenças
    query_sites = "SELECT DISTINCT Sites FROM Site"
    sites = pd.read_sql(query_sites, conn)['Sites'].tolist()
    presenca_opcoes = pd.read_sql("SELECT DISTINCT Presenca FROM Presenca", conn)['Presenca'].tolist()

    # Captura os valores dos filtros e salva na sessão
    selected_site = request.form.get("site") or session.get('selected_site')
    selected_empresa = request.form.get("empresa") or session.get('selected_empresa')
    if selected_site:
        session['selected_site'] = selected_site
    if selected_empresa:
        session['selected_empresa'] = selected_empresa

    # Inicializa variáveis
    nomes = []
    nomes_desativados = []
    empresas = []
    empresas_inativas = []
    registros_mes_atual = []
    siteempresa_id = None

    # Obter ano e mês atuais
    current_year = datetime.now().year
    current_month = datetime.now().month  # Obtem o mês como inteiro

    if selected_site:
        empresas = get_empresas(get_site_id(selected_site))
        empresas_inativas = get_empresas_inativas(get_site_id(selected_site))

        if selected_empresa:
            site_id = get_site_id(selected_site)
            empresa_id = get_empresa_id(selected_empresa, empresas)
            siteempresa_id = get_siteempresa_id(site_id, empresa_id)

            if site_id and empresa_id:
                # Chama a função fetch_registros_mes para buscar os registros do mês atual
                registros_mes_atual = fetch_registros_mes(site_id, empresa_id, current_month, current_year)

                # Obtém nomes ativos e desativados
                if siteempresa_id:
                    nomes = get_nomes(siteempresa_id, ativos=True)
                    nomes_desativados = get_nomes(siteempresa_id, ativos=False)
    
    current_month = datetime.now().strftime("%m")

    conn.close()
    # Garante que o gráfico seja gerado apenas se houver registros
    if registros_mes_atual:
        df = pd.DataFrame(registros_mes_atual, columns=['Nome', 'Presenca', 'Data'])
        
        # Converte a coluna 'Data' para datetime
        df['Data'] = pd.to_datetime(df['Data'], format='%Y-%m-%d %H:%M:%S')

        # Gera uma lista contínua de datas entre o menor e o maior valor de data
        min_data = df['Data'].min()
        max_data = df['Data'].max()
        datas_continuas = pd.date_range(min_data, max_data).to_list()

        # Cria um DataFrame contínuo para evitar lacunas no gráfico
        nomes_unicos = df['Nome'].unique()
        df_continuo = pd.MultiIndex.from_product([nomes_unicos, datas_continuas], names=['Nome', 'Data']).to_frame(index=False)

        # Converte ambas as colunas 'Data' para datetime para garantir a compatibilidade no merge
        df_continuo['Data'] = pd.to_datetime(df_continuo['Data'])
        df['Data'] = pd.to_datetime(df['Data'])

        # Faz o merge do DataFrame original com o DataFrame contínuo
        df_merge = pd.merge(df_continuo, df, on=['Nome', 'Data'], how='left')

        # Preenche valores ausentes com "invisível" para manter a estrutura do gráfico
        df_merge['Presenca'].fillna('invisível', inplace=True)

        # Criando o gráfico de dispersão
        fig_dispersao = go.Figure()

        for presenca, info in color_marker_map.items():
            df_tipo = df_merge[df_merge['Presenca'].str.upper() == presenca]
            if not df_tipo.empty:
                fig_dispersao.add_trace(go.Scatter(
                    x=df_tipo['Data'],
                    y=df_tipo['Nome'],
                    mode='markers',
                    marker=dict(color=info['cor'], symbol=info['marker'], size=10),
                    name=presenca
                ))

        # Adicionar pontos invisíveis para garantir espaçamento correto no gráfico
        df_invisivel = df_merge[df_merge['Presenca'] == 'invisível']
        fig_dispersao.add_trace(go.Scatter(
            x=df_invisivel['Data'],
            y=df_invisivel['Nome'],
            mode='markers',
            marker=dict(color='rgba(0,0,0,0)', size=10),
            name='invisível',
            showlegend=False
        ))

        # Layout do gráfico igual ao da página principal
        fig_dispersao.update_layout(
            title={
                #'text': "Gráfico de Dispersão de Presenças",
                'x': 0.5,
                'xanchor': 'center',
                'yanchor': 'top',
                'font': {'size': 24}
            },
            xaxis=dict(
                showgrid=False,
                gridcolor='lightgray',
                tickformat='%d/%m/%Y',
                dtick=86400000
            ),
            yaxis=dict(showgrid=False, gridcolor='lightgray'),
            font=dict(color='#000000'),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            hovermode='closest'
        )

        # Converte o gráfico para JSON e passa para o template
        scatter_chart_data = json.dumps(fig_dispersao, cls=plotly.utils.PlotlyJSONEncoder)
    else:
        scatter_chart_data = None  # Se não houver dados, evita erro


    # Renderiza o template HTML
    return render_template(
        "adicionar_presenca.html",
        sites=sites,
        dias = [str(i).zfill(2) for i in range(1, 32)],
        empresas=[e[1] for e in empresas],
        empresas_inativas=[e[1] for e in empresas_inativas],
        selected_site=selected_site,
        selected_empresa=selected_empresa,
        siteempresa_id=siteempresa_id,
        nomes=nomes,
        nomes_desativados=nomes_desativados,
        presenca_opcoes=presenca_opcoes,
        current_month=current_month,
        current_year=current_year,
        last_year=current_year - 1,
        registros_mes_atual=registros_mes_atual,
        meses_dict=meses_dict,
        color_marker_map=color_marker_map,
        scatter_chart_data=scatter_chart_data,
    )



# __________________ ROTAS PARA FLUXO _________________
@app.route('/reativar-nome', methods=['POST'])
def reativar_nome():
    nome_desativado = request.form.get("nome_desativado").strip()
    siteempresa_id = request.form.get("siteempresa_id")

    print(f"Nome desativado: {nome_desativado}, SiteEmpresa ID: {siteempresa_id}")

    if not nome_desativado:
        flash("Nenhum nome selecionado para reativar!", "error")
        return redirect(url_for('adiciona_presenca'))

    try:
        conn = get_db_connection()
        if not conn:
            flash("Erro ao conectar ao banco de dados.", "error")
            return redirect(url_for('adiciona_presenca'))
        
        cursor = conn.cursor()
        cursor.execute("UPDATE Nome SET Ativo = True WHERE Nome = ? AND id_SiteEmpresa = ?",
                       (nome_desativado, siteempresa_id))
        conn.commit()
        conn.close()
        flash(f"Nome {nome_desativado} reativado com sucesso!", "success")
    except Exception as e:
        print(f"Erro ao reativar nome: {e}")
        flash(f"Erro ao reativar nome: {str(e)}", "error")

    return redirect(url_for('adiciona_presenca'))

@app.route('/inativar-nome', methods=['POST'])
def inativar_nome():
    nome_ativo = request.form.get("nome_ativo").strip()
    siteempresa_id = request.form.get("siteempresa_id")

    print(f"Nome ativo: {nome_ativo}, SiteEmpresa ID: {siteempresa_id}")

    if not nome_ativo:
        flash("Nenhum nome selecionado para desativar!", "error")
        return redirect(url_for('adiciona_presenca'))

    try:
        conn = get_db_connection()
        if not conn:
            flash("Erro ao conectar ao banco de dados.", "error")
            return redirect(url_for('adiciona_presenca'))

        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM Nome WHERE id_SiteEmpresa = ? AND Ativo = True", (siteempresa_id,))
        num_nomes_ativos = cursor.fetchone()[0]

        if num_nomes_ativos <= 1:
            flash("Não é possível desativar o último nome ativo. Pelo menos um nome deve permanecer ativo.", "error")
            conn.close()
            return redirect(url_for('adiciona_presenca'))

        cursor.execute("UPDATE Nome SET Ativo = False WHERE Nome = ? AND id_SiteEmpresa = ?",
                       (nome_ativo, siteempresa_id))
        conn.commit()
        conn.close()

        flash(f"Nome {nome_ativo} desativado com sucesso!", "success")
    except Exception as e:
        print(f"Erro ao desativar nome: {e}")
        flash(f"Erro ao desativar nome: {str(e)}", "error")

    return redirect(url_for('adiciona_presenca'))

@app.route('/presenca', methods=['POST'])
def controlar_presenca():
    nomes = request.form.getlist('nomes')
    tipo_presenca = request.form.get('presenca')
    dia = request.form.get('dia')
    mes = request.form.get('mes')
    ano = request.form.get('ano')
    siteempresa_id = request.form.get('siteempresa_id')
    action_type = request.form.get('action_type')

    if not nomes or not dia or not mes or not ano:
        flash("Por favor, selecione todos os campos: Nomes, Dia, Mês e Ano.", "error")
        return redirect(url_for('adiciona_presenca'))

    try:
        ano = int(ano)
        mes = int(mes)
        dia = int(dia)
        
        if mes == 2 and dia in [29, 30, 31] and not calendar.isleap(ano):
            flash(f"O ano {ano} não é bissexto, portanto, fevereiro não pode ter mais de 28 dias.", "error")
            return redirect(url_for('adiciona_presenca'))
        
        data_selecionada = datetime(ano, mes, dia)
        dia_semana = data_selecionada.weekday()

        if dia_semana >= 5:
            flash("Não é permitido adicionar presença em sábados ou domingos.", "error")
            return redirect(url_for('adiciona_presenca'))

        conn = get_db_connection()
        if not conn:
            flash("Erro ao conectar ao banco de dados.", "error")
            return redirect(url_for('adiciona_presenca'))

        cursor = conn.cursor()
        nomes_adicionados = []
        nomes_atualizados = []

        if action_type == 'adicionar':
            for nome in nomes:
                cursor.execute("SELECT id_Nomes FROM Nome WHERE Nome = ? AND id_SiteEmpresa = ?", (nome, siteempresa_id))
                id_nome = cursor.fetchone()[0]

                cursor.execute("""
                    SELECT id_Controle FROM Controle 
                    WHERE id_Nome = ? AND Data = ? AND id_SiteEmpresa = ?
                """, (id_nome, data_selecionada, siteempresa_id))
                id_controle = cursor.fetchone()

                cursor.execute("SELECT id_Presenca FROM Presenca WHERE Presenca = ?", (tipo_presenca,))
                id_presenca = cursor.fetchone()[0]

                if id_controle:
                    cursor.execute("""
                        UPDATE Controle 
                        SET id_Presenca = ?
                        WHERE id_Controle = ?
                    """, (id_presenca, id_controle[0]))
                    nomes_atualizados.append(nome)
                else:
                    cursor.execute("""
                        INSERT INTO Controle (id_Nome, id_Presenca, Data, id_SiteEmpresa)
                        VALUES (?, ?, ?, ?)
                    """, (id_nome, id_presenca, data_selecionada, siteempresa_id))
                    nomes_adicionados.append(nome)

            conn.commit()

            if nomes_adicionados:
                flash(f"Presença adicionada com sucesso para os nomes: {', '.join(nomes_adicionados)} na data {data_selecionada.strftime('%d/%m/%Y')}", "success")
            if nomes_atualizados:
                flash(f"Presença atualizada com sucesso para os nomes: {', '.join(nomes_atualizados)} na data {data_selecionada.strftime('%d/%m/%Y')}", "warning")

        elif action_type == 'remover':
            for nome in nomes:
                cursor.execute("SELECT id_Nomes FROM Nome WHERE Nome = ? AND id_SiteEmpresa = ?", (nome, siteempresa_id))
                id_nome = cursor.fetchone()[0]

                cursor.execute("""
                    SELECT id_Controle FROM Controle
                    WHERE id_Nome = ? AND Data = ? AND id_SiteEmpresa = ?
                """, (id_nome, data_selecionada, siteempresa_id))
                id_controle = cursor.fetchone()

                if id_controle:
                    cursor.execute("DELETE FROM Controle WHERE id_Controle = ?", (id_controle[0],))
                else:
                    flash(f"Não foi encontrado registro de presença para {nome} na data {data_selecionada.strftime('%d/%m/%Y')}.", "error")

            conn.commit()
            flash(f"Presença removida para os nomes: {', '.join(nomes)} na data {data_selecionada.strftime('%d/%m/%Y')}", "remover")

        conn.close()
    except pyodbc.Error as e:
        flash(f"Erro ao realizar a ação de presença: {e}", "error")

    return redirect(url_for('adiciona_presenca'))

@app.route('/adicionar-nome', methods=['POST'])
def adicionar_nome():
    novo_nome = request.form.get("novo_nome")
    siteempresa_id = request.form.get("siteempresa_id")

    if not novo_nome or not siteempresa_id:
        flash("Por favor, preencha todos os campos.", "error")
        return redirect(url_for('adiciona_presenca'))

    try:
        conn = get_db_connection()
        if not conn:
            flash("Erro ao conectar ao banco de dados.", "error")
            return redirect(url_for('adiciona_presenca'))

        novo_nome = novo_nome.strip().title()
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM Nome WHERE Nome = ? AND id_SiteEmpresa = ?", (novo_nome, siteempresa_id))
        existe_nome = cursor.fetchone()[0]

        if existe_nome > 0:
            flash(f"O nome '{novo_nome}' já existe na tabela.", "warning")
            conn.close()
            return redirect(url_for('adiciona_presenca'))

        cursor.execute("SELECT MAX(id_Nomes) FROM Nome")
        ultimo_id = cursor.fetchone()[0] or 0
        novo_id = ultimo_id + 1

        cursor.execute(
            """
            INSERT INTO Nome (id_Nomes, id_SiteEmpresa, Nome, Ativo)
            VALUES (?, ?, ?, ?)
            """, (novo_id, siteempresa_id, novo_nome, True))

        conn.commit()
        conn.close()
        flash(f"Nome '{novo_nome}' adicionado com sucesso!", "success")
    except Exception as e:
        flash(f"Erro ao adicionar nome: {e}", "error")

    return redirect(url_for('adiciona_presenca'))


@app.route('/adicionar-empresa', methods=['POST'])
def adicionar_empresa():
    site_nome = request.form.get("site") or session.get('selected_site')
    nova_empresa = request.form.get("nova_empresa").strip()

    if not site_nome or not nova_empresa:
        flash("Por favor, preencha todos os campos.", "error")
        return redirect(url_for('adiciona_presenca'))

    try:
        conn = get_db_connection()
        if not conn:
            flash("Erro ao conectar ao banco de dados.", "error")
            return redirect(url_for('adiciona_presenca'))

        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM Empresa WHERE Empresas = ?", (nova_empresa,))
        existe_empresa = cursor.fetchone()[0]

        if existe_empresa > 0:
            flash(f"A empresa '{nova_empresa}' já existe.", "warning")
            conn.close()
            return redirect(url_for('adiciona_presenca'))

        cursor.execute("SELECT MAX(id_Empresa) FROM Empresa")
        ultimo_id_empresa = cursor.fetchone()[0] or 0
        novo_id_empresa = ultimo_id_empresa + 1

        cursor.execute(
            """
            INSERT INTO Empresa (id_Empresa, Empresas)
            VALUES (?, ?)
            """, (novo_id_empresa, nova_empresa))

        site_id = get_site_id(site_nome)
        if not site_id:
            flash("Site não encontrado.", "error")
            conn.close()
            return redirect(url_for('adiciona_presenca'))

        cursor.execute(
            """
            INSERT INTO Site_Empresa (id_Sites, id_Empresas, Ativo)
            VALUES (?, ?, ?)
            """, (site_id, novo_id_empresa, True))

        conn.commit()
        conn.close()
        flash(f"Empresa '{nova_empresa}' adicionada com sucesso ao site '{site_nome}'!", "success")
    except Exception as e:
        flash(f"Erro ao adicionar empresa: {str(e)}", "error")

    return redirect(url_for('adiciona_presenca'))


@app.route('/desativar-empresa', methods=['POST'])
def desativar_empresa():
    empresa_ativa = request.form.get("empresa_ativa")

    if not empresa_ativa:
        flash("Nenhuma empresa selecionada para desativar.", "error")
        return redirect(url_for('adiciona_presenca'))

    try:
        conn = get_db_connection()
        if not conn:
            flash("Erro ao conectar ao banco de dados.", "error")
            return redirect(url_for('adiciona_presenca'))

        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM Site_Empresa WHERE Ativo = True")
        num_empresas_ativas = cursor.fetchone()[0]

        if num_empresas_ativas <= 1:
            flash("Não é possível desativar todas as empresas. Pelo menos uma empresa deve estar ativa.", "error")
            conn.close()
            return redirect(url_for('adiciona_presenca'))

        empresa_selecionada = session.get('selected_empresa')
        if empresa_selecionada == empresa_ativa:
            flash(f"A empresa '{empresa_ativa}' está em uso e não pode ser desativada.", "error")
            conn.close()
            return redirect(url_for('adiciona_presenca'))

        cursor.execute("SELECT id_Empresa FROM Empresa WHERE Empresas = ?", (empresa_ativa,))
        id_empresa = cursor.fetchone()[0]

        cursor.execute("UPDATE Site_Empresa SET Ativo = False WHERE id_Empresas = ?", (id_empresa,))
        conn.commit()
        conn.close()
        flash(f"Empresa '{empresa_ativa}' desativada com sucesso!", "success")
    except Exception as e:
        flash(f"Erro ao desativar a empresa: {str(e)}", "error")

    return redirect(url_for('adiciona_presenca'))


@app.route('/ativar-empresa', methods=['POST'])
def ativar_empresa():
    empresa_inativa = request.form.get("empresa_inativa")

    if not empresa_inativa:
        flash("Nenhuma empresa selecionada para ativar.", "error")
        return redirect(url_for('adiciona_presenca'))

    try:
        conn = get_db_connection()
        if not conn:
            flash("Erro ao conectar ao banco de dados.", "error")
            return redirect(url_for('adiciona_presenca'))

        cursor = conn.cursor()
        cursor.execute("SELECT id_Empresa FROM Empresa WHERE Empresas = ?", (empresa_inativa,))
        id_empresa = cursor.fetchone()[0]

        cursor.execute("UPDATE Site_Empresa SET Ativo = True WHERE id_Empresas = ?", (id_empresa,))
        conn.commit()
        conn.close()
        flash(f"Empresa '{empresa_inativa}' ativada com sucesso!", "success")
    except Exception as e:
        flash(f"Erro ao ativar a empresa: {str(e)}", "error")

    return redirect(url_for('adiciona_presenca'))

@app.route('/programa-ferias', methods=['POST'])
def programa_ferias():
    nome = request.form.get('nome_ativo')
    data_inicio = request.form.get('data_inicio')
    data_fim = request.form.get('data_fim')
    siteempresa_id = request.form.get('siteempresa_id')

    if not nome or not data_inicio or not data_fim:
        flash("Por favor, preencha todos os campos.", "error")
        return redirect(url_for('adiciona_presenca'))

    try:
        conn = get_db_connection()
        if not conn:
            flash("Erro ao conectar ao banco de dados.", "error")
            return redirect(url_for('adiciona_presenca'))

        data_inicio = datetime.strptime(data_inicio, '%Y-%m-%d')
        data_fim = datetime.strptime(data_fim, '%Y-%m-%d')

        if data_inicio > data_fim:
            flash("A data de início não pode ser maior que a data de fim.", "error")
            conn.close()
            return redirect(url_for('adiciona_presenca'))

        cursor = conn.cursor()
        cursor.execute("SELECT id_Nomes FROM Nome WHERE Nome = ? AND id_SiteEmpresa = ?", (nome, siteempresa_id))
        id_nome_result = cursor.fetchone()

        if id_nome_result is None:
            flash(f"Nome '{nome}' não encontrado para o site/empresa selecionado.", "error")
            conn.close()
            return redirect(url_for('adiciona_presenca'))

        id_nome = id_nome_result[0]
        cursor.execute("SELECT id_Presenca FROM Presenca WHERE Presenca = 'FÉRIAS'")
        id_presenca_result = cursor.fetchone()

        if id_presenca_result is None:
            flash("Tipo de presença 'FÉRIAS' não encontrado.", "error")
            conn.close()
            return redirect(url_for('adiciona_presenca'))

        id_presenca = id_presenca_result[0]
        cursor.execute(
            """
            SELECT COUNT(*) FROM Controle 
            WHERE id_Nome = ? AND id_Presenca = ? AND id_SiteEmpresa = ?
            """, (id_nome, id_presenca, siteempresa_id))
        total_dias_ferias = cursor.fetchone()[0]

        dias_programados = (data_fim - data_inicio).days + 1

        if total_dias_ferias + dias_programados > 30:
            flash(f"O nome '{nome}' já tem {total_dias_ferias} dias de férias programados. Com esses novos {dias_programados} dias, o total excede o limite de 30 dias.", "error")
            conn.close()
            return redirect(url_for('adiciona_presenca'))

        current_date = data_inicio
        while current_date <= data_fim:
            cursor.execute(
                """
                INSERT INTO Controle (id_Nome, id_Presenca, Data, id_SiteEmpresa)
                VALUES (?, ?, ?, ?)
                """, (id_nome, id_presenca, current_date, siteempresa_id))
            current_date += timedelta(days=1)

        conn.commit()
        conn.close()
        flash(f"Férias programadas com sucesso para {nome} de {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}", "success")
    except Exception as e:
        flash(f"Erro ao programar férias: {e}", "error")

    return redirect(url_for('adiciona_presenca'))

@app.route('/desprogramar-ferias', methods=['POST'])
def desprogramar_ferias():
    nome = request.form.get('nome_ativo')
    data_inicio = request.form.get('data_inicio')
    data_fim = request.form.get('data_fim')
    siteempresa_id = request.form.get('siteempresa_id')

    if not nome or not data_inicio or not data_fim:
        flash("Por favor, preencha todos os campos.", "error")
        return redirect(url_for('adiciona_presenca'))

    try:
        conn = get_db_connection()
        if not conn:
            flash("Erro ao conectar ao banco de dados.", "error")
            return redirect(url_for('adiciona_presenca'))

        data_inicio = datetime.strptime(data_inicio, '%Y-%m-%d')
        data_fim = datetime.strptime(data_fim, '%Y-%m-%d')

        if data_inicio > data_fim:
            flash("A data de início não pode ser maior que a data de fim.", "error")
            conn.close()
            return redirect(url_for('adiciona_presenca'))

        cursor = conn.cursor()
        cursor.execute("SELECT id_Nomes FROM Nome WHERE Nome = ? AND id_SiteEmpresa = ?", (nome, siteempresa_id))
        id_nome_result = cursor.fetchone()

        if id_nome_result is None:
            flash(f"Nome '{nome}' não encontrado para o site/empresa selecionado.", "error")
            conn.close()
            return redirect(url_for('adiciona_presenca'))

        id_nome = id_nome_result[0]
        cursor.execute("SELECT id_Presenca FROM Presenca WHERE Presenca = 'FÉRIAS'")
        id_presenca_result = cursor.fetchone()

        if id_presenca_result is None:
            flash("Tipo de presença 'FÉRIAS' não encontrado.", "error")
            conn.close()
            return redirect(url_for('adiciona_presenca'))

        id_presenca = id_presenca_result[0]
        current_date = data_inicio
        while current_date <= data_fim:
            cursor.execute(
                """
                DELETE FROM Controle 
                WHERE id_Nome = ? AND id_Presenca = ? AND Data = ? AND id_SiteEmpresa = ?
                """, (id_nome, id_presenca, current_date, siteempresa_id))
            current_date += timedelta(days=1)

        conn.commit()
        conn.close()
        flash(f"Férias desprogramadas com sucesso para {nome} de {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}", "success")
    except Exception as e:
        flash(f"Erro ao desprogramar férias: {e}", "error")

    return redirect(url_for('adiciona_presenca'))



#* -------------------------
#*         CHATBOT
#* -------------------------
def identificar_nomes(tokens):
    """Verifica se uma palavra é um nome baseado nos nomes cadastrados no banco de dados."""
    conn = get_db_connection()
    if not conn:
        return None
    cursor = conn.cursor()
    query = "SELECT Nome FROM Nome"
    cursor.execute(query)
    nomes_cadastrados = {row[0].lower() for row in cursor.fetchall()}
    conn.close()

    for token in tokens:
        if token in nomes_cadastrados:
            return token.capitalize()  # Retorna o nome encontrado
    return None

# Função para identificar a intenção e responder corretamente
def identificar_pergunta(user_input):
    user_input = user_input.lower().strip()
    melhor_score = 0.0
    melhor_correspondencia = None

    for categoria, subcategorias in perguntas_respostas.items():
        for subcategoria, dados in subcategorias.items():
            if "perguntas" in dados:
                for pergunta in dados["perguntas"]:
                    score = difflib.SequenceMatcher(None, user_input, pergunta).ratio()
                    if score > melhor_score:
                        melhor_score = score
                        melhor_correspondencia = {"tipo": "ajuda", "mensagem": dados["resposta"]}

    if melhor_score > 0.6:
        return melhor_correspondencia  # Retorna a resposta se encontrou uma boa correspondência

    return None  # Retorna None se não encontrou uma pergunta relevante

def processar_mensagem(mensagem):
    """
    Processa a mensagem do usuário e extrai:
      - nome_input: nome da pessoa
      - periodo: lista de meses/anos (máximo 2 elementos: mês e ano)
      - tipo_frequencia: tipo de presença (convertido para singular)
    """
    resultado_pergunta = identificar_pergunta(mensagem)
    if resultado_pergunta:
        return {"tipo": "ajuda", "mensagem": resultado_pergunta["mensagem"]}

    tokens = word_tokenize(mensagem.lower())  # Tokeniza a mensagem
    tokens = [t for t in tokens if t.isalnum()]  # Remove pontuações desnecessárias
    periodo = []
    nome_input = None
    tipo_frequencia = None

    # Padrões de períodos de data
    padrao_ano = re.compile(r"\b\d{4}\b")  # Apenas ANO (ex.: "2024")
    padrao_mes_ano_barra = re.compile(r"(\b\d{1,2})/(\d{4}\b)")  # "MM/YYYY"
    padrao_mes_ano_de = re.compile(r"(\b[a-zA-Z]{3,9}|\b\d{1,2})\s+de\s+(\d{4}\b)")  # "setembro de 2024" ou "09 de 2024"

    # ✅ Verifica "MM de YYYY" ou "Mês Abreviado de YYYY" (ex.: "set de 2024", "setembro de 2024", "09 de 2024")
    match_mes_ano_de = padrao_mes_ano_de.search(mensagem)
    if match_mes_ano_de:
        mes_texto = match_mes_ano_de.group(1).lower()
        ano = match_mes_ano_de.group(2)

        mes_numerico = meses_map.get(mes_texto, mes_texto.zfill(2))  # Converte para número se for nome
        if mes_numerico not in periodo:
            periodo.append(mes_numerico)
        if ano not in periodo:
            periodo.append(ano)

    # ✅ Verifica "MM/YYYY" (ex.: "09/2024")
    match_mes_ano_barra = padrao_mes_ano_barra.search(mensagem)
    if match_mes_ano_barra:
        mes = match_mes_ano_barra.group(1).zfill(2)
        ano = match_mes_ano_barra.group(2)
        if mes not in periodo:
            periodo.append(mes)
        if ano not in periodo:
            periodo.append(ano)

    # ✅ Verifica apenas o ANO (ex.: "2024"), e só adiciona se não houver um ano na lista
    match_ano = padrao_ano.search(mensagem)
    if match_ano:
        ano_encontrado = match_ano.group(0)
        if len(periodo) == 0:  # Se não há mês, adiciona apenas o ano
            periodo.append(ano_encontrado)
        elif len(periodo) == 1 and periodo[0].isdigit() and len(periodo[0]) == 2:
            periodo.append(ano_encontrado)  # Se já tem um mês, adiciona o ano

    # ✅ Remove duplicatas e mantém no máximo 2 elementos (mês e ano)
    periodo = list(dict.fromkeys(periodo))  # Remove duplicatas preservando a ordem correta
    periodo = periodo[:2]  # Mantém apenas os dois primeiros valores (mês e ano)

    for palavra in tokens:
        if palavra in meses_map and meses_map[palavra] not in periodo:
            periodo.append(meses_map[palavra])
            continue
        if palavra.isdigit() and len(palavra) == 4 and palavra not in periodo:  # Evita duplicação de anos
            periodo.append(palavra)
            continue

        # ✅ Identifica tipo de frequência (ex.: "faltas" → "falta")
        if palavra in frequencia_plural_para_singular:
            tipo_frequencia = frequencia_plural_para_singular[palavra]
        elif palavra in frequencia_plural_para_singular.values():
            tipo_frequencia = palavra

    # 🔹 Busca o nome diretamente no banco de dados:
    nome_input = identificar_nomes(tokens)

    # print(f"🔍 Processamento: Nome={nome_input}, Periodo={periodo}, Tipo_Frequencia={tipo_frequencia}")

    return {
        "nome_input": nome_input,
        "periodo": periodo if len(periodo) > 0 else None,
        "tipo_frequencia": tipo_frequencia,
        "tipo": "QUERY",
    }


# ******************
#*  ROTAS DO CHATBOT
#*  - QUERY DO BANCO
# ******************
def listar_nomes_disponiveis():
    try:
        conn = get_db_connection()
        if not conn:
            return
        query = "SELECT Nome.Nome FROM Nome"
        nomes_disponiveis = pd.read_sql(query, conn)
        conn.close()

        if not nomes_disponiveis.empty:
            # Converte o DataFrame em tabela HTML com classe para facilitar a estilização
            html_table = nomes_disponiveis.to_html(classes="content-table", index=False, border=0)
            return f"<h3>Nomes presentes na consulta</h3>" + html_table
        
        else:
                return f"<h3>Nomes presentes na consulta</h3><p>Nenhum dado encontrado.</p>"
    except Exception as e:
        return f"<h3>Nomes presentes na consulta</h3><p>Erro na consulta: {str(e)}</p>"

def executar_consulta(query, params, titulo):
    """
    Executa a query SQL usando os parâmetros fornecidos, converte o resultado em um DataFrame
    e retorna uma string HTML contendo o título e uma tabela formatada.
    """
    try:
        conn = get_db_connection()
        # Usando pandas para ler a query – observe que o pyodbc aceita parâmetro via "params"
        df = pd.read_sql(query, conn, params=params)
        conn.close()
        if not df.empty:
            # Converte o DataFrame em tabela HTML com classe para facilitar a estilização
            html_table = df.to_html(classes="content-table", index=False, border=0)
            return f"<h3>{titulo}</h3>" + html_table
        else:
            return f"<h3>{titulo}</h3><p>Nenhum dado encontrado.</p>"
    except Exception as e:
        return f"<h3>{titulo}</h3><p>Erro na consulta: {str(e)}</p>"

def consulta_presencas(nome_input, periodo, tipo_frequencia):
    periodo_formatado = f"{periodo[0]}/{periodo[1]}" if len(periodo) == 2 else f"%/{periodo[0]}"
    query = """
        SELECT Nome.Nome, Presenca.Presenca, FORMAT(Controle.Data, 'mm/yyyy') AS Mes,
               COUNT(Controle.id_Controle) AS TotalPresencas
        FROM (Controle
        INNER JOIN Nome ON Controle.id_Nome = Nome.id_Nomes)
        INNER JOIN Presenca ON Controle.id_Presenca = Presenca.id_Presenca
        WHERE Presenca.Presenca = ? AND Nome.Nome = ? AND FORMAT(Controle.Data, 'mm/yyyy') LIKE ?
        GROUP BY Nome.Nome, Presenca.Presenca, FORMAT(Controle.Data, 'mm/yyyy')
        ORDER BY COUNT(Controle.id_Controle) DESC;
    """
    return executar_consulta(query, [tipo_frequencia, nome_input, periodo_formatado], "Quantidade de Presenças")

def consulta_presenca_por_nome(nome_input, periodo):
    """
    Consulta as presenças de um nome específico dentro de um período de tempo.
    """
    query = """
        SELECT Nome.Nome, Presenca.Presenca, FORMAT(Controle.Data, 'mm/yyyy') AS MesAno,
               COUNT(Controle.id_Controle) AS TotalPresencas
        FROM (Controle
        INNER JOIN Nome ON Controle.id_Nome = Nome.id_Nomes)
        INNER JOIN Presenca ON Controle.id_Presenca = Presenca.id_Presenca
        WHERE Nome.Nome = ?
    """
    params = [nome_input]

    # Se o usuário passou um mês e um ano (ex.: ["09", "2024"])
    if len(periodo) == 2:
        query += " AND FORMAT(Controle.Data, 'mm/yyyy') = ?"
        params.append(f"{periodo[0]}/{periodo[1]}")

    # Se o usuário passou apenas o ano (ex.: ["2024"])
    elif len(periodo) == 1:
        query += " AND YEAR(Controle.Data) = ?"
        params.append(periodo[0])

    query += """
        GROUP BY Nome.Nome, Presenca.Presenca, FORMAT(Controle.Data, 'mm/yyyy')
        ORDER BY FORMAT(Controle.Data, 'mm/yyyy') ASC;
    """
    
    return executar_consulta(query, params, "Contagem de Presenças por Nome")

def consulta_nome_mais_presencas(tipo_frequencia, periodo=None):
    query = """
        SELECT Nome.Nome, Presenca.Presenca, FORMAT(Controle.Data, 'mm/yyyy') AS MesAno,
               COUNT(Controle.id_Controle) AS TotalPresencas
        FROM ((Controle
        INNER JOIN Nome ON Controle.id_Nome = Nome.id_Nomes)
        INNER JOIN Presenca ON Controle.id_Presenca = Presenca.id_Presenca)
        WHERE Presenca.Presenca = ?
    """
    params = [tipo_frequencia]
    
    if periodo:
        if len(periodo) == 2:  # Caso seja MÊS e ANO
            mes = meses_map.get(periodo[0].lower(), periodo[0])  # Converte "setembro" -> "09"
            mes = mes.zfill(2) if mes.isdigit() else mes  # Garante formato "08"
            query += " AND FORMAT(Controle.Data, 'mm/yyyy') = ?"
            params.append(f"{mes}/{periodo[1]}")
        else:  # Caso seja apenas o ANO
            query += " AND YEAR(Controle.Data) = ?"
            params.append(periodo[0])

    # Finaliza a query
    query += """
        GROUP BY Nome.Nome, Presenca.Presenca, FORMAT(Controle.Data, 'mm/yyyy')
        ORDER BY COUNT(Controle.id_Controle) DESC;
    """
    
    return executar_consulta(query, params, f"Mais presenças de '{tipo_frequencia}' no período {periodo if periodo else 'geral'}")

def consulta_nome_mais_presenca_msg(tipo_frequencia, periodo=None):
    query = """
    SELECT TOP 1 Nome.Nome, Presenca.Presenca, FORMAT(Controle.Data, 'mm/yyyy') AS MesAno,
               COUNT(Controle.id_Controle) AS TotalPresencas
    FROM (Controle
    INNER JOIN Nome ON Controle.id_Nome = Nome.id_Nomes)
    INNER JOIN Presenca ON Controle.id_Presenca = Presenca.id_Presenca  
    WHERE Presenca.Presenca = ?
    """

    params = [tipo_frequencia]

    # Ajusta a query caso período seja especificado
    if periodo:
        if len(periodo) == 2:
            mes = meses_map.get(periodo[0].lower(), periodo[0])  
            mes = mes.zfill(2) if mes.isdigit() else mes
            query += " AND FORMAT(Controle.Data, 'mm/yyyy') = ?"
            params.append(f"{mes}/{periodo[1]}")
        else:
            query += " AND FORMAT(Controle.Data, 'yyyy') = ?"
            params.append(periodo[0])

    query += """
        GROUP BY Nome.Nome, Presenca.Presenca, FORMAT(Controle.Data, 'mm/yyyy')
        ORDER BY COUNT(Controle.id_Controle) DESC;
    """
    
    conn = get_db_connection()
    # Usando pandas para ler a query – observe que o pyodbc aceita parâmetro via "params"
    df = pd.read_sql(query, conn, params=params)
    conn.close()
    if not df.empty:
       teste = df.iloc[0].tolist()
       return teste
    else:
        return None
    
def consulta_por_presenca_e_periodo(tipo_frequencia, periodo):
    query = """
        SELECT Nome.Nome, Presenca.Presenca, FORMAT(Controle.Data, 'yyyy') AS Ano,
               COUNT(Controle.id_Controle) AS TotalPresencas
        FROM (Controle
        INNER JOIN Nome ON Controle.id_Nome = Nome.id_Nomes)
        INNER JOIN Presenca ON Controle.id_Presenca = Presenca.id_Presenca
        WHERE Presenca.Presenca = ? AND YEAR(Controle.Data) = ?
        GROUP BY Nome.Nome, Presenca.Presenca, FORMAT(Controle.Data, 'yyyy')
        ORDER BY FORMAT(Controle.Data, 'yyyy') ASC;
    """
    return executar_consulta(query, [tipo_frequencia, periodo[0]], "Presenças por Tipo e Período")

def consulta_todas_presencas(nome_input):
    query = """
        SELECT Nome.Nome, Presenca.Presenca, FORMAT(Controle.Data, 'mm/yyyy') AS MesAno,
               COUNT(Controle.id_Controle) AS TotalPresencas
        FROM (Controle
        INNER JOIN Nome ON Controle.id_Nome = Nome.id_Nomes)
        INNER JOIN Presenca ON Controle.id_Presenca = Presenca.id_Presenca
        WHERE Nome.Nome = ?
        GROUP BY Nome.Nome, Presenca.Presenca, FORMAT(Controle.Data, 'mm/yyyy')
        ORDER BY FORMAT(Controle.Data, 'mm/yyyy') ASC;
    """
    return executar_consulta(query, [nome_input], f"Todas as presenças de {nome_input}")

def consulta_todas_presencas_periodo(periodo):
    query = """
        SELECT Nome.Nome, Presenca.Presenca, FORMAT(Controle.Data, 'mm/yyyy') AS MesAno,
               COUNT(Controle.id_Controle) AS TotalPresencas
        FROM (Controle
        INNER JOIN Nome ON Controle.id_Nome = Nome.id_Nomes)
        INNER JOIN Presenca ON Controle.id_Presenca = Presenca.id_Presenca
        WHERE FORMAT(Controle.Data, 'yyyy') = ?
    """
    # Garante que estamos pegando o ano corretamente
    ano = periodo[0] if len(periodo) == 1 or len(periodo[0]) == 4 else periodo[1]
    params = [ano]
    
    # Se o usuário também forneceu um mês, adiciona esse filtro
    if len(periodo) == 2:
        mes = periodo[1] if periodo[0] == ano else periodo[0]
        mes_numero = meses_map.get(mes.lower(), mes)  # Converte "outubro" → "10" se necessário
        mes_numero = mes_numero.zfill(2) if mes_numero.isdigit() else mes_numero  # Garante formato "08"
        query += " AND FORMAT(Controle.Data, 'mm/yyyy') = ?"
        params.append(f"{mes_numero}/{ano}")
    
    query += """
        GROUP BY Nome.Nome, Presenca.Presenca, FORMAT(Controle.Data, 'mm/yyyy')
        ORDER BY FORMAT(Controle.Data, 'mm/yyyy') ASC;
    """
    
    return executar_consulta(query, params, f"Todas as presenças registradas no período {periodo}")



# CHATBOT_FUNCTION
@app.route("/chatbot", methods=["GET", "POST"])
def chatbot():
    """ Rota para processar mensagens do chat """
    if request.method == "POST":
        dados = request.get_json()
        mensagem_usuario = dados.get("mensagem", "").strip()
        
        respostas = []  # Lista que armazenará cada mensagem a ser retornada

        # Verifica se o usuário quer encerrar a conversa
        if mensagem_usuario.lower() in ["sair", "exit", "quit", "tchau", "até logo", "adeus", "encerrar"]:
            respostas.append({
                "tipo": "text",
                "mensagem": "Até logo! Foi um prazer ajudar você."
            })
            return jsonify({"respostas": respostas})
        
        match = difflib.get_close_matches(mensagem_usuario.lower(), saudacoes_validas, n=1)
        if match:
            respostas.append({
                "tipo": "text",
                "mensagem": "Olá! Como posso ajudar você hoje?"
            })
            return jsonify({"respostas": respostas})
        
        # Se o usuário pedir a lista de nomes disponíveis
        if difflib.get_close_matches(mensagem_usuario, LISTAGEM_NOMES, n=1, cutoff=0.6):
            lista_nomes = listar_nomes_disponiveis()  # Supondo que essa função retorne uma string ou lista formatada
            respostas.append({
                "tipo": "text",
                "mensagem": "Claro! Aqui estão os nomes disponíveis:"
            })
            respostas.append({
                "tipo": "table",
                "mensagem": lista_nomes
            })

        # Processa a mensagem para extrair os parâmetros
        processamento = processar_mensagem(mensagem_usuario)
        
        # Caso seja uma pergunta de ajuda, retorna a resposta imediatamente e encerra
        if processamento["tipo"] == "ajuda":
            return jsonify({"respostas": [{"tipo": "text", "mensagem": processamento["mensagem"]}]})

        # Dependendo dos dados extraídos, chame a consulta apropriada.
        if processamento["nome_input"] and processamento["tipo_frequencia"] and processamento["periodo"]:
            resultado = consulta_presencas(
                processamento["nome_input"],
                processamento["periodo"],
                processamento["tipo_frequencia"]
            )
            respostas.append({
                "tipo": "table",
                "mensagem": resultado
            })
        elif processamento["tipo_frequencia"]:
            resultado = consulta_nome_mais_presencas(
                processamento["tipo_frequencia"],
                processamento["periodo"]
            )
            respostas.append({
                "tipo": "table",
                "mensagem": resultado
            })
            resultado_dois = consulta_nome_mais_presenca_msg(
                processamento["tipo_frequencia"],
                processamento["periodo"]
            )

            respostas.append({
                "tipo": "text",
                "mensagem": f'{resultado_dois[0]} teve mais {resultado_dois[1]} no mês de {resultado_dois[2]} com um total de {resultado_dois[3]} presenças.'
            })
        elif processamento["nome_input"] and processamento["periodo"]:
            resultado = consulta_presenca_por_nome(
                processamento["nome_input"],
                processamento["periodo"]
            )
            respostas.append({
                "tipo": "table",
                "mensagem": resultado
            })
        elif processamento["tipo_frequencia"] and processamento["periodo"]:
            resultado = consulta_por_presenca_e_periodo(
                processamento["tipo_frequencia"],
                processamento["periodo"]
            )
            respostas.append({
                "tipo": "table",
                "mensagem": resultado
            })
        elif processamento["nome_input"] and not processamento["tipo_frequencia"] and not processamento["periodo"]:
            resultado = consulta_todas_presencas(processamento["nome_input"])
            respostas.append({
                "tipo": "table",
                "mensagem": resultado
            })
        elif processamento["periodo"] and not processamento["nome_input"] and not processamento["tipo_frequencia"]:
            resultado = consulta_todas_presencas_periodo(processamento["periodo"])
            respostas.append({
                "tipo": "table",
                "mensagem": resultado
            })
        
        # Caso nenhuma condição tenha sido satisfeita, envie uma mensagem padrão.
        if not respostas:
            respostas.append({
                "tipo": "text",
                "mensagem": "Desculpe, não entendi sua solicitação. Poderia reformular?"
            })
        
        return jsonify({"respostas": respostas})

    # Para requisição GET, envia uma mensagem padrão.
    return jsonify({"respostas": [{"tipo": "text", "mensagem": "Bem-vindo ao chatbot! Envie uma mensagem para começar."}]})




if __name__ == "__main__":
    verificar_e_instalar_nltk()
    rprint('\n\t   :snake: [b]DASHBOARD - CONTROLE DE FREQUENCIA[/] :snake:')
    rprint('[d]_______________________________________________________________[/]\n')
    rprint('Voce consegue visualizar o seu Dashboard atraves da URL\n')
    rprint('[d]URL :[/] [blink b] http://127.0.0.1:5000 [/]')
    rprint('\t\t[blue d] ↑ Copie e cole a url em qualquer navegador![/]')
    print('\n')
    rprint(':clock10: A[blue blink] URL [/]SÓ FUNCIONARA SE MANTER O EXECUTAVEL [b]ABERTO[/]')
    print('\n')
    rprint('[on red] Press CTRL+C para fechar[/]\n')
    app.run()