from pathlib import Path
from docxtpl import DocxTemplate
import pandas as pd

# 1. Configurações de caminhos (ajuste conforme a sua pasta)
diretorio_base = Path("./Documentos/arthur/certificados")
caminho_excel = diretorio_base / "dados.xlsx"
caminho_modelo = diretorio_base / "modelo_certificado.docx"
pasta_saida = diretorio_base / "certificados_gerados"

# Cria a pasta de saída caso ela não exista
pasta_saida.mkdir(parents=True, exist_ok=True)

# 2. Lê a planilha do Excel (dtype=str evita que CPFs percam zeros à esquerda)
df = pd.read_excel(caminho_excel, dtype=str)from pathlib import Path
from docxtpl import DocxTemplate
import pandas as pd

# 1. Configurações de caminhos (ajuste conforme a sua pasta)
diretorio_base = Path("./arquivos")
caminho_excel = diretorio_base / "dados.xlsx"
caminho_modelo = diretorio_base / "modelo_certificado.docx"
pasta_saida = diretorio_base / "certificados_gerados"

# Cria a pasta de saída caso ela não exista
pasta_saida.mkdir(parents=True, exist_ok=True)

# 2. Lê a planilha do Excel (dtype=str evita que CPFs percam zeros à esquerda)
df = pd.read_excel(caminho_excel, dtype=str)

# Remove possíveis espaços extras nos nomes das colunas
df.columns = df.columns.str.strip()

# 3. Itera sobre cada linha da planilha
for index, linha in df.iterrows():
    nome = str(linha["Nome"]).strip()
    cpf = str(linha["CPF"]).strip()

    # Pula linhas vazias, caso existam
    if nome == "nan" or not nome:
        continue

    # Carrega o modelo a cada iteração para não sobrescrever os anteriores
    doc = DocxTemplate(caminho_modelo)

    # Dicionário com as variáveis que estão no Word
    contexto = {"nome": nome, "cpf": cpf}

    # Substitui as tags no documento
    doc.render(contexto)

    # Limpa caracteres inválidos para nome de arquivo e salva
    nome_arquivo = f"Certificado - {nome}.docx".replace("/", "-")
    caminho_final = pasta_saida / nome_arquivo

    doc.save(caminho_final)
    print(f"Gerado: {nome_arquivo}")

print("\nProcesso concluído!")

# Remove possíveis espaços extras nos nomes das colunas
df.columns = df.columns.str.strip()

# 3. Itera sobre cada linha da planilha
for index, linha in df.iterrows():
    nome = str(linha["Nome"]).strip()
    cpf = str(linha["CPF"]).strip()

    # Pula linhas vazias, caso existam
    if nome == "nan" or not nome:
        continue

    # Carrega o modelo a cada iteração para não sobrescrever os anteriores
    doc = DocxTemplate(caminho_modelo)

    # Dicionário com as variáveis que estão no Word
    contexto = {"nome": nome, "cpf": cpf}

    # Substitui as tags no documento
    doc.render(contexto)

    # Limpa caracteres inválidos para nome de arquivo e salva
    nome_arquivo = f"Certificado - {nome}.docx".replace("/", "-")
    caminho_final = pasta_saida / nome_arquivo

    doc.save(caminho_final)
    print(f"Gerado: {nome_arquivo}")

print("\nProcesso concluído!")
