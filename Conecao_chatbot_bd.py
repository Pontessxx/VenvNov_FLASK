import pandas as pd
import pyodbc
import spacy
from bs4 import BeautifulSoup
import requests
import warnings
import difflib


warnings.filterwarnings("ignore", category=UserWarning)

# Carregar o modelo SpaCy
nlp = spacy.load("pt_core_news_sm")

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

import difflib

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



def listar_nomes_disponiveis():
    conn = get_db_connection()
    if not conn:
        return

    cursor = conn.cursor()
    query = "SELECT Nome.Nome FROM Nome"
    cursor.execute(query)
    nomes_disponiveis = [row[0] for row in cursor.fetchall()]
    conn.close()

    if nomes_disponiveis:
        print("\n🔹 **Lista de Nomes Disponíveis no Banco** 🔹")
        for nome in nomes_disponiveis:
            print(f"- {nome}")
    else:
        print("\n⚠ Nenhum nome encontrado no banco de dados.")


# 📌 **Função para identificar a intenção e responder corretamente**
def identificar_pergunta(user_input):
    user_input = user_input.lower().strip()
    
    melhor_correspondencia = None
    melhor_score = 0.0

    # Percorre todas as categorias (adicionar, remover, filtrar)
    for categoria, subcategorias in perguntas_respostas.items():
        # Se for uma categoria sem subcategorias (ex.: "filtrar")
        if isinstance(subcategorias, dict) and "perguntas" in subcategorias:
            for pergunta in subcategorias["perguntas"]:
                score = difflib.SequenceMatcher(None, user_input, pergunta).ratio()
                if score > melhor_score:
                    melhor_score = score
                    melhor_correspondencia = {
                        "tipo": "ajuda",
                        "mensagem": subcategorias["resposta"]
                    }
        
        # Se for uma categoria com subcategorias (ex.: "adicionar", "remover")
        elif isinstance(subcategorias, dict):
            for subcategoria, dados in subcategorias.items():
                if "perguntas" in dados:  # Garantir que a chave "perguntas" existe antes de acessar
                    for pergunta in dados["perguntas"]:
                        score = difflib.SequenceMatcher(None, user_input, pergunta).ratio()
                        if score > melhor_score:
                            melhor_score = score
                            melhor_correspondencia = {
                                "tipo": "ajuda",
                                "mensagem": dados["resposta"]
                            }

    # Se encontrou uma correspondência com alto grau de similaridade, retorna a resposta
    if melhor_score > 0.6:  # Ajuste fino para precisão, pode ser aumentado para evitar confusões
        return melhor_correspondencia

    return {
        "tipo": "erro",
        "mensagem": "Desculpe, não entendi sua dúvida. Poderia reformular?"
    }



# Função para processar a frase do usuário e extrair informações relevantes
def process_user_input(user_input):
    doc = nlp(user_input)
    nome_input = None
    periodo = []
    tipo_frequencia = None

    # 🟢 **Primeiro, verifica se o usuário fez uma pergunta sobre o sistema**
    resultado_pergunta = identificar_pergunta(user_input)
    if resultado_pergunta:
        return resultado_pergunta

    for token in doc:
        palavra = token.text.lower()

        if palavra in meses_map:
            periodo.append(meses_map[palavra])
            continue

        if palavra.isdigit() and 1 <= int(palavra) <= 12:
            periodo.append(palavra.zfill(2))
            continue

        if palavra.isdigit() and len(palavra) == 4:
            periodo.append(palavra)
            continue

        if palavra in frequencia_plural_para_singular:
            tipo_frequencia = frequencia_plural_para_singular[palavra]
        elif palavra in frequencia_plural_para_singular.values():
            tipo_frequencia = palavra

        if token.pos_ == "PROPN" and palavra not in meses_map.values():
            nome_input = token.text

    return {"nome_input": nome_input, "periodo": periodo, "tipo_frequencia": tipo_frequencia}


# Função para executar consultas e exibir os resultados
def executar_consulta(query, params, titulo):
    conn = get_db_connection()
    if not conn:
        return

    df = pd.read_sql(query, conn, params=params)
    conn.close()

    print(f"\n🔹 {titulo} 🔹")
    if not df.empty:
        print(df.to_string(index=False))
    else:
        print("\n⚠ Nenhum resultado encontrado.")

# 📌 **Consulta 1**: Quantidade de presenças filtrando por Mês/Ano/Nome
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
    executar_consulta(query, [tipo_frequencia, nome_input, periodo_formatado], "Quantidade de Presenças")

# 📌 **Consulta 2**: Nome com mais presenças de um tipo específico
def consulta_nome_mais_presencas(tipo_frequencia, periodo=None):
    query = """
        SELECT Nome.Nome, Presenca.Presenca, FORMAT(Controle.Data, 'mm/yyyy') AS MesAno,
               COUNT(Controle.id_Controle) AS TotalPresencas
        FROM (Controle
        INNER JOIN Nome ON Controle.id_Nome = Nome.id_Nomes)
        INNER JOIN Presenca ON Controle.id_Presenca = Presenca.id_Presenca
        WHERE Presenca.Presenca = ?
    """

    params = [tipo_frequencia]

    # Se houver um período especificado (mês e ano), adiciona o filtro de período
    if periodo:
        if len(periodo) == 2:
            mes = meses_map.get(periodo[0].lower(), periodo[0])  # Converte "agosto" -> "08"
            mes = mes.zfill(2) if mes.isdigit() else mes  # Garante "08" para meses numéricos
            query += " AND FORMAT(Controle.Data, 'mm/yyyy') = ?"
            params.append(f"{mes}/{periodo[1]}")
        else:
            query += " AND FORMAT(Controle.Data, 'yyyy') = ?"
            params.append(periodo[0])

    # Finaliza a query corretamente
    query += """
        GROUP BY Nome.Nome, Presenca.Presenca, FORMAT(Controle.Data, 'mm/yyyy')
        ORDER BY COUNT(Controle.id_Controle) DESC;
    """

    executar_consulta(query, params, f"Mais presenças de '{tipo_frequencia}' no período {periodo if periodo else 'geral'}")

# 📌 **Consulta 3**: Contagem de cada tipo de presença para um nome específico dentro de um ano
def consulta_presenca_por_nome(nome_input, periodo):
    
    
    query = """
        SELECT Nome.Nome, Presenca.Presenca, FORMAT(Controle.Data, 'mm/yyyy') AS MesAno,
               COUNT(Controle.id_Controle) AS TotalPresencas
        FROM (Controle
        INNER JOIN Nome ON Controle.id_Nome = Nome.id_Nomes)
        INNER JOIN Presenca ON Controle.id_Presenca = Presenca.id_Presenca
        WHERE Nome.Nome = ? AND FORMAT(Controle.Data, 'yyyy') = ?
        GROUP BY Nome.Nome, Presenca.Presenca, FORMAT(Controle.Data, 'mm/yyyy')
        ORDER BY FORMAT(Controle.Data, 'mm/yyyy') ASC;
    """
    executar_consulta(query, [nome_input, periodo[0]], "Contagem de Presenças por Nome")

# 📌 **Consulta 4**: Filtrar por período e tipo de presença (retorna nomes)
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
    executar_consulta(query, [tipo_frequencia, periodo[0]], "Presenças por Tipo e Período")

# 📌 **Consulta 5: para obter todas as presenças de um nome sem filtro de período**
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
    executar_consulta(query, [nome_input], f"Todas as presenças de {nome_input}")

# 📌 **Consulta 6**: Obter todas as presenças filtradas por Mês e Ano (se disponível)
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

    # Se o usuário também forneceu um mês, adicionar esse filtro
    if len(periodo) == 2:
        mes = periodo[1] if periodo[0] == ano else periodo[0]  # Garante que mes está correto
        mes_numero = meses_map.get(mes.lower(), mes)  # Converte "outubro" → "10" se necessário
        mes_numero = mes_numero.zfill(2) if mes_numero.isdigit() else mes_numero  # Garante formato "08"
        query += " AND FORMAT(Controle.Data, 'mm/yyyy') = ?"
        params.append(f"{mes_numero}/{ano}")  # Corrigindo o formato da consulta

    # Finalizar a query corretamente
    query += """
        GROUP BY Nome.Nome, Presenca.Presenca, FORMAT(Controle.Data, 'mm/yyyy')
        ORDER BY FORMAT(Controle.Data, 'mm/yyyy') ASC;
    """
    
    executar_consulta(query, params, f"Todas as presenças registradas no período {periodo}")



# Função principal do chatbot
def chatbot():
    print("Olá! Bem-vindo ao meu assistente virtual. 😊")
    user_name = input("Insira seu nome: ").strip()
    print(f"\nOlá, {user_name}! O que gostaria de fazer hoje?")
    
    while True:
        user_input = input(f"{user_name}: ").strip()
        
        # Se o usuário perguntar pelos nomes disponíveis, chamamos a função de listagem
        if "nomes disponivel" in user_input.lower() or "lista de nomes" in user_input.lower():
            listar_nomes_disponiveis()
            continue

        
        if user_input.lower() in ["sair", "exit", "quit", "tchau", "até logo", "adeus", "encerrar"]:
            print(f"Até logo, {user_name}! Foi um prazer ajudar você.")
            break
        
        resultado = process_user_input(user_input)
        print("\n🔍 Analisando sua solicitação...")
        
        if resultado.get("tipo") == "ajuda":
            print(f"\n🤖 {resultado['mensagem']}")
            continue  

        if resultado["nome_input"] and resultado["tipo_frequencia"] and resultado["periodo"]:
            consulta_presencas(resultado["nome_input"], resultado["periodo"], resultado["tipo_frequencia"])
        elif resultado["tipo_frequencia"]:
            consulta_nome_mais_presencas(resultado["tipo_frequencia"], resultado["periodo"])
        elif resultado["nome_input"] and resultado["periodo"]:
            consulta_presenca_por_nome(resultado["nome_input"], resultado["periodo"])
        elif resultado["tipo_frequencia"] and resultado["periodo"]:
            consulta_por_presenca_e_periodo(resultado["tipo_frequencia"], resultado["periodo"])
        elif resultado["nome_input"] and not resultado["tipo_frequencia"] and not resultado["periodo"]:
            consulta_todas_presencas(resultado["nome_input"])
        elif resultado["periodo"] and not resultado["nome_input"] and not resultado["tipo_frequencia"]:
            consulta_todas_presencas_periodo(resultado["periodo"])


        print("-" * 40)

# Iniciar o chatbot
chatbot()
