

O objetivo foi analisar uma base de leads e identificar quais características apresentam maior
chance de conversão em clientes, apoiando decisões do time de marketing.

## Estrutura do projeto

- 01_Analise_Dados.xlsx  
  Arquivo principal de análise. Contém a base consolidada de leads e as análises por:
  origem, mercado, porte e objetivo, além do racional utilizado.

- Apresentacao_Desafio.pptx  
  Apresentação final com os principais insights, perfil ideal de lead e recomendações
  para o time de marketing. Pode ser entendida sem explicação oral.

- analise.py  
  Script em Python utilizado para reproduzir a análise de forma programática e gerar
  tabelas consolidadas a partir do Excel.

- app.py  
  Dashboard interativo desenvolvido em Streamlit para visualização dos resultados
  (opcional, utilizado como complemento visual).

- README.md  
  Documento explicativo com instruções e visão geral do projeto.


## Ferramentas utilizadas

- Excel (análise e tabelas dinâmicas)
- Python (pandas)
- Streamlit (dashboard interativo)
- PowerPoint (apresentação dos resultados)


## Fonte dos dados

Os códigos em Python utilizam o arquivo `01_Analise_Dados.xlsx`, localizado na pasta do projeto.
A aba principal utilizada é `BASE_ANALISE`, onde os dados já estão consolidados.


## Como executar os códigos 

1. Ter Python instalado
2. Instalar as bibliotecas necessárias:
   pip install pandas streamlit

3. Executar o script de análise:
   python analise.py

4. Executar o dashboard:
   streamlit run app.py

## Ajuste do caminho do arquivo Excel

Os códigos em Python utilizam o arquivo de dados localizado no caminho abaixo:

C:\Users\joaov\Documents\DESAFIO ESTAGIO\01_Analise_Dados.xlsx

Caso o projeto seja executado em outra máquina, é necessário **ajustar o caminho do arquivo**
para o local onde o Excel estiver salvo no computador do usuário.

Esse ajuste deve ser feito diretamente no código Python, na linha onde o arquivo é carregado,
alterando apenas o caminho para o novo local do arquivo.


## Observações

- Métricas de custo, CAC e ROI são simuladas e utilizadas apenas para análise estratégica.
- A recomendação final do desafio é baseada principalmente na taxa de conversão
  e no volume de leads, evitando conclusões com amostras muito pequenas.
