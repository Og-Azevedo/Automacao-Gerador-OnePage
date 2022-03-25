# Gerador automático de relatórios

**Desafio:**

Uma franquia possui 25 lojas no Brasil todo. O seu sistema financeiro gera o final de cada dia um relatório unificado contendo os dados de vendas referente a todas as 25 lojas. Por todos os dados estarem aglomerados dentro de um só arquivo o acompanhamento, inteligência, agilidade na tomada de decisão ficam comprometidos.
Esse trabalho atualmente é feito manualmente por vários funcionários, o que resulta em um processo extremamente ineficaz e propenso ao erro humano. Além disso, é uma enorme alocação de recursos da empresa. 

**Objetivo:**

Criar um programa que receba apenas um arquivo em excel execute todos os dias as seguintes ações:
  - Gerar planilhas segmentadas para cada loja;
  - Criar Backup;
  - Tratar dados;
  - Gerar um 'One Page' sintetizando e sinalizando se as metas foram batidas;
  - Enviar planilhas para cada gerente responsável por cada loja;
  - Enviar rankings enviar para diretoria;

**Baseado no desafio e no objetivo apresentado eu desenvolvi um programa em python que possui todas essas funcionalidades seguintes:**

**Funcionalidades:**

**1- Criar Backup:**
  - Receber uma planilha unificada em .xlsx contendo histórico de vendas de 25 lojas;
  - Gerar um arquivo no formato xlsx para cada loja e nome no formato: "DIA_MÊS_NOME DA LOJA.xlsx"
  - Criar uma pasta de backup, e em seguida uma pasta para cada loja.
  - Salvar cada planilha em sua pasta correspondente no formato xlsx.

 **2- Tratar dados:**
   - Verificar se cada loja bateu as metas:
     - Faturamento -> Meta Ano: 1.650.000 / Meta Dia: 1000
     - Diversidade de Produtos: -> Meta Ano: 120 / Meta Dia: 4
     - Ticket Médio por Venda: -> Meta Ano: 500 / Meta Dia: 500
   - Gerar 'One Page': Baseado nesse cruzamento de dados deve ser gerado um texto que sintetize e sinalize visualmente para o gerente da loja se as metas(anuais e diárias) daquela filial foram batidas. Segue exemplo abaixo:
   
![image](https://user-images.githubusercontent.com/14143617/160150738-dbe74793-7fd4-4061-8c4c-adfcf4c1351e.png) ![image](https://user-images.githubusercontent.com/14143617/160151164-b9e62417-4a03-47e4-aba5-c129d8187004.png)

**3- Enviar dados para gerentes:**
  - Buscar na planilha 'Base de Dados/'Emails.xlsx'' o email do gerente responsável por cada loja;
  - Enviar para cada gerente o 'One Page' e a planilha (.xlsx) referente ao desempenho do último dia;
 
**4- Criar rankings/OnePage e enviar para diretoria**
  - Usando a planilha unificada com todos os históricos o programa vai gerar: dois rankings(.xlsx) e um 'One Page':
    - Ranking do faturamento anual das lojas (decrescente);
    - Ranking  do faturamento das lojas referente ao último dia registrado(decrescente);
    - 'One Page' baseado no faturamento informando a melhor e pior loja do ano e do último dia registrado. Segue exemplo:
      - A melhor loja do Ranking diário foi: Salvador Shopping com faturamento de R$3950
      - A pior loja do Ranking diário foi : Shopping Ibirapuera com faturamento de R$118

      - A melhor loja do Ranking anual foi: Iguatemi Campinas com faturamento de R$3950
      - A pior loja do Ranking anual foi : Shopping Morumbi com faturamento de R$118
  - Buscar na planilha 'Base de Dados/'Emails.xlsx'' o email da diretoria;
  - Enviar os dois rankings e o 'One Page' para o email da diretoria;

