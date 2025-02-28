import shutil
import nltk

# Caminho padrão onde o NLTK baixa os arquivos
nltk_data_path = nltk.data.find('.').path

# Remove a pasta nltk_data e todo o seu conteúdo
shutil.rmtree(nltk_data_path, ignore_errors=True)

print("Todos os dados do NLTK foram removidos!")
