# 🤖 Robô de Faturamento - Correios MRV

Este projeto é uma solução completa de RPA (Robotic Process Automation) desenvolvida em Python para automatizar o fluxo de faturamento, conciliação de rateios e lançamento de notas fiscais dos Correios no portal da MRV.

## 🎯 Funcionalidades

O sistema possui uma Interface Gráfica (GUI) dividida em 3 etapas independentes:

1. **Gerar Rascunhos (Outlook):** Varre a rede corporativa em busca dos extratos do mês atual e gera automaticamente rascunhos de e-mail no Outlook com os anexos e textos padronizados para cada regional.
2. **Gerar Planilha de Rateio:** Lê a planilha original dos Correios e a planilha de resposta da regional, realiza a conciliação de valores, rateia encargos/descontos matematicamente e gera o arquivo final `RATEIO PAG.xlsx`.
3. **Lançar Nota (Portal MRV):** Robô web (Selenium) que lê o boleto em PDF, extrai CNPJ, datas e valores, cruza com a base de regras de negócio, faz login no portal da MRV, preenche todos os formulários, anexa a planilha silenciosamente e finaliza o lançamento.

## 📁 Estrutura de Pastas e Arquivos Necessários

Para que o robô funcione corretamente, a seguinte estrutura de arquivos deve ser respeitada:

| Arquivo / Pasta | Descrição |
| --- | --- |
| `app.py` | Arquivo principal que contém a interface e a lógica do robô. |
| `dados_puxados_preenchimento.xlsx` | Base de regras (De-Para) com CNPJs, IDs e Códigos de Material. Deve ficar na mesma pasta do `app.py`. |
| `testar_edicao/` | Pasta usada na **Etapa 2**. Deve conter a planilha original dos Correios (ex: `1234567.xlsx`) e a planilha da regional (com   `Rateio Recebido` no nome). |
| `exemplos/` | Pasta usada na **Etapa 3**. Deve conter apenas **UM** boleto em `.pdf` e a planilha `RATEIO PAG.xlsx` (gerada na Etapa 2). |

## 🚀 Como Executar

1. Certifique-se de ter o Python 3 instalado na máquina.
2. Instale as dependências do projeto executando no terminal:

pip install pandas openpyxl selenium PyPDF2 pywin32


Execute o aplicativo: python app.py

A interface gráfica será aberta. Clique nos botões na ordem do processo. O console integrado exibirá os logs em tempo real.

🛠️ Tecnologias Utilizadas

Tkinter & Threading: Interface gráfica e processamento assíncrono.
Selenium WebDriver: Automação web e injeção de JavaScript para alta performance.
Pandas & Openpyxl: Manipulação e conciliação de dados em Excel.
PyPDF2 & Regex: Extração inteligente de dados não estruturados de PDFs.
Win32com: Integração nativa com o Microsoft Outlook.