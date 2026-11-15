from datetime import datetime
from pathlib import Path
import subprocess
from docxtpl import DocxTemplate
import pandas as pd

# 1. Configurações de caminhos
diretorio_base = Path(".")
caminho_excel = diretorio_base / "copia com dados.xlsx"
caminho_modelo = diretorio_base / "Certificado_SOS_Dom_Almir_Modelo_Editavel.docx"
pasta_saida = diretorio_base / "certificados_gerados"

# Cria a pasta de saída caso ela não exista
pasta_saida.mkdir(parents=True, exist_ok=True)

# 2. Lê a planilha do Excel (dtype=str evita que CPFs percam zeros à esquerda)
df = pd.read_excel(caminho_excel, dtype=str)

# Remove possíveis espaços extras nos nomes das colunas
df.columns = df.columns.str.strip()

# 3. Pega a data atual do sistema e formata como DD/MM/YYYY
data_atual = datetime.now().strftime("%d/%m/%Y")

# 4. Itera sobre cada linha da planilha
for index, linha in df.iterrows():
    nome = str(linha["Nome"]).strip()
    cpf = str(linha["CPF"]).strip()

    # Pula linhas caso o Nome esteja vazio/nan
    if nome == "nan" or not nome:
        continue

    # Pula linhas caso o CPF esteja vazio, nan ou nulo
    if cpf == "nan" or not cpf or cpf.lower() == "none":
        print(f"Pulado (CPF em branco): {nome}")
        continue

    # Carrega o modelo a cada iteração para não sobrescrever os anteriores
    doc = DocxTemplate(caminho_modelo)

    # Dicionário com as variáveis que estão no Word
    contexto = {"nome": nome, "cpf": cpf, "data_emissao": data_atual}

    # Substitui as tags no documento
    doc.render(contexto)

    # Nome base para os arquivos (substituindo barras para evitar erro no Linux)
    nome_limpo = f"Certificado - {nome}".replace("/", "-")
    caminho_docx = pasta_saida / f"{nome_limpo}.docx"

    # Salva o arquivo temporário em .docx
    doc.save(caminho_docx)

    # Converte o .docx gerado para PDF usando o LibreOffice via terminal do Python
    try:
        subprocess.run(
            [
                "soffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(pasta_saida),
                str(caminho_docx),
            ],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )

        # (Opcional) Se quiser apagar o arquivo .docx e deixar apenas o PDF na pasta,
        # descomente a linha abaixo removendo a cerquilha (#):
        # caminho_docx.unlink()

        print(f"Gerado PDF: {nome_limpo}.pdf")

    except Exception as e:
        print(f"Erro ao converter o certificado de {nome}: {e}")

print("\nProcesso concluído!")
