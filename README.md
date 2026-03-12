📘 Processamento Automático Setores CellDowntime (BA)
Este projeto automatiza o download, extração, tratamento e consolidação do arquivo Setores_CellDowntime_HMM, filtrando dados para o estado da Bahia e consolidando informações por ERB.
O fluxo é totalmente configurável via arquivo .env, permitindo fácil manutenção e atualização.

📌 Funcionalidades
✔ Download automático do arquivo ZIP
Baixa o arquivo diretamente da URL configurada no .env.
✔ Extração do arquivo Excel interno
Extrai apenas o arquivo desejado de dentro do ZIP, também configurado no .env.
✔ Detecção automática do Excel

Aba com dados
Linha de cabeçalho
Coluna de data mais recente

✔ Processamento dos dados
Inclui:

Filtro por UF = BA
Filtro de valores ≥ 500 na coluna de data mais recente

✔ Consolidação por ERB
Para cada ERB:

Mantém apenas uma linha
Cria coluna QTD_ERB com a quantidade de linhas originais
Soma os valores da coluna de data mais recente
Mantém demais colunas fixas (REGIONAL, MUNICIPIO, SITE etc.)

✔ Merge opcional com uma nova planilha
Se existir nova_planilha.xlsx, o sistema adiciona novas colunas via join por ERB.
✔ Geração da planilha final
Arquivo salvo automaticamente em:
./saida/Setores_Celldowntime_BA.xlsx

📂 Estrutura do Projeto
/
├── script.py
├── .env
├── .gitignore
├── nova_planilha.xlsx # opcional
└── saida/
└── Setores_Celldowntime_BA.xlsx
