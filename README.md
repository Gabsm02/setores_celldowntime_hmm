📘 – Processamento Automático Setores CellDowntime BA
Este projeto automatiza o processo de download, extração e tratamento do arquivo Setores_CellDowntime_HMM, filtrando informações da Bahia (BA) e consolidando dados por ERB.
O fluxo é totalmente configurável via arquivo .env, tornando o processo flexível e seguro.

🚀 Funcionalidades Principais
✔ 1. Carregamento de configurações via .env
O script utiliza o arquivo .env para definir:

URL do arquivo ZIP
Nome do arquivo a extrair de dentro do ZIP

Isso evita hardcode no código e permite trocas rápidas sem editar o script.

✔ 2. Download automático do ZIP
O script baixa automaticamente o arquivo ZIP remoto usando a URL configurada:
URL_ZIP=<link configurado no .env>


✔ 3. Extração do arquivo interno
De dentro do ZIP, apenas o arquivo especificado é extraído:
ALVO_NO_ZIP=<arquivo .xlsx do .env>


✔ 4. Detecção inteligente do Excel
O script identifica automaticamente:

Aba com dados
Linha de cabeçalho
Coluna de data mais recente

Tudo isso sem precisar editar código.

✔ 5. Filtragem dos dados
Após carregar os dados, o script aplica:

Filtro UF = BA
Filtro de valores ≥ 500 na coluna de data mais recente


✔ 6. Consolidação por ERB
Para cada ERB:

Cria a coluna QTD_ERB indicando quantas linhas pertenciam àquela ERB
Mantém uma única linha por ERB
Mantém o primeiro valor da coluna de data mais recente
Mantém todas as outras colunas fixas (REGIONAL, MUNICIPIO etc.)


✔ 7. Merge opcional com nova planilha
Se existir nova_planilha.xlsx, o script realiza:

Merge via ERB
Mantendo os dados novos ao lado dos dados originais


✔ 8. Geração da planilha final
A planilha final é salva em:
/saida/Setores_Celldowntime_BA.xlsx


📁 Estrutura do Projeto
/
├── script.py
├── .env
├── .gitignore
├── nova_planilha.xlsx        # opcional
└── saida/
    └── Setores_Celldowntime_BA.xlsx


⚙️ Arquivo .env
Crie um arquivo .env com o seguinte conteúdo:
Plain Textenv não tem suporte total. O realce de sintaxe é baseado em Plain Text.URL_ZIP=https://maestro.vivo.com.br/movel/downloads/setores_celldowntime_HMM.zipALVO_NO_ZIP=Setores_CellDowntime_HMM_NE_0.xlsxMostrar mais linhas

🧹 Arquivo .gitignore
O repositório utiliza um .gitignore para evitar envio de arquivos sensíveis e de saída:
.env
saida/
*.xlsx
__pycache__/
venv/


📦 Dependências
Instale as dependências:
Shellpip install pandas openpyxl urllib3 python-dotenv``Mostrar mais linhas

▶️ Como Executar
Simples:
Shellpython script.pyMostrar mais linhas

📄 Saída Final
A planilha final contém:

REGIONAL
UF
MUNICIPIO
CN
SITE
TECNOLOGIA
ERB
SETOR
COLUNA DE DATA MAIS RECENTE
QTD_ERB (quantas linhas originais existiam da ERB)

Há apenas uma linha por ERB.

🔧 Customização
Tudo pode ser ajustado via:

.env
Nome dos arquivos
Regras de merge
Regras de agregação
