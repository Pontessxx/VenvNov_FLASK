import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk import pos_tag

# Lista de pacotes necessários e seus respectivos caminhos dentro do nltk_data
required_packages = {
    "punkt": "tokenizers/punkt",
    "stopwords": "corpora/stopwords",
    "averaged_perceptron_tagger": "taggers/averaged_perceptron_tagger"
}

# Verifica se os pacotes já estão instalados e baixa apenas os ausentes
for package, path in required_packages.items():
    try:
        nltk.data.find(path)
        print(f"✔ {package} já está instalado.")
    except LookupError:
        print(f"⬇ Baixando {package}...")
        nltk.download(package)

# Exemplo de uso com um texto em português
texto = "Este é um teste simples para tokenização e remoção de stopwords."

# Tokenização
tokens = word_tokenize(texto, language="portuguese")

# Removendo stopwords em português
stop_words = set(stopwords.words("portuguese"))
tokens_filtrados = [word for word in tokens if word.lower() not in stop_words]

# POS Tagging (NLTK não tem suporte oficial para PT-BR, então usa modelo em inglês como fallback)
pos_tags = pos_tag(tokens_filtrados)

# Exibir resultados
print("\nTokens Filtrados:", tokens_filtrados)
print("\nPOS Tags:", pos_tags)
