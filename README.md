# 🎓 Gerador Automático de Certificados em Python

Script automatizado desenvolvido em Python para gerar certificados personalizados em formato `.docx` em lote, a partir de uma planilha do Excel (`.xlsx`) e de um modelo predefinido no Word.

## 🚀 Funcionalidades
* **Leitura automatizada do Excel:** Extrai nomes e CPFs preservando zeros à esquerda (essencial para CPFs válidos).
* **Validação de dados:** Pula automaticamente linhas com CPFs em branco ou nulos, evitando certificados corrompidos.
* **Preenchimento de variáveis:** Substitui tags dinâmicas (`{{ nome }}`, `{{ cpf }}`, `{{ data_emissao }}`) no modelo do Word.
* **Data dinâmica:** Insere automaticamente a data de emissão atualizada no formato `DD/MM/YYYY`.
* **Preservação de formatação:** Utiliza a engine `docxtpl` para manter todas as fontes, cores e o layout original do documento Word.

---

## 🛠️ Tecnologias Utilizadas
* [Python 3](https://www.python.org/)
* [Pandas](https://pandas.pydata.org/) (Manipulação da planilha Excel)
* [OpenPyXL](https://openpyxl.readthedocs.io/) (Leitura de arquivos `.xlsx`)
* [Python-Docx-Template](https://docxtpl.readthedocs.io/) (Renderização do modelo `.docx`)

---

## 📂 Estrutura do Projeto
```text
certificados/
│
├── .venv/                              # Ambiente virtual Python (ignorado pelo git)
├── certificados_gerados/               # Pasta de saída com os certificados prontos
├── .gitignore                          # Arquivos ignorados pelo controle de versão
├── copia com dados.xlsx                # Planilha com os dados dos participantes
├── Certificado_Modelo_Editavel.docx    # Modelo base do certificado
└── gerar_certificados.py               # Script principal de automação
