import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk import pos_tag

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
            print(f"✔ {package} já está instalado.")
        except LookupError:
            print(f"⬇ Baixando {package}...")
            nltk.download(package)

# Chama a função para garantir que os pacotes estão disponíveis
verificar_e_instalar_nltk()

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
