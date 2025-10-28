# Projeto de Automação e Análise de Dados

Este é um projeto desenvolvido durante um curso de Python. O script automatiza o cálculo de indicadores, a criação de relatórios e o envio de e-mails para diferentes níveis hierárquicos (gerentes de loja e diretoria) em uma rede de varejo.

---

## Objetivo

Imagine que você trabalha em uma grande rede de lojas de roupa com 25 lojas espalhadas pelo Brasil. Todo dia, a equipe de análise de dados precisa calcular os "One Pages" e enviá-los para o gerente de cada loja, junto com as informações usadas no cálculo.

Um **One Page** é um resumo direto ao ponto, usado pela gerência para saber os principais indicadores da loja, comparar resultados e verificar o cumprimento de metas.

O objetivo deste script é **automatizar 100% desse processo**: desde a leitura dos dados brutos até o envio final dos e-mails.

## Funcionalidades Principais

O script executa as seguintes tarefas:

1.  **Leitura e Consolidação:** Carrega e trata as bases de dados de vendas, lojas e e-mails.
2.  **Cálculo de Indicadores:** Calcula três indicadores-chave (Faturamento, Diversidade de Produtos, Ticket Médio) para cada loja, tanto no dia quanto no acumulado do ano.
3.  **Segregação de Dados:** Filtra a base de vendas mestra e cria planilhas individuais contendo as vendas de *apenas* uma loja.
4.  **Envio para Gerentes:** Envia um e-mail individual para cada gerente de loja contendo:
    * Um **"One Page"** no corpo do e-mail, com os indicadores da sua loja e um sinal visual (🟢/🔴) de cumprimento de meta.
    * Em anexo, a planilha Excel com as vendas detalhadas **apenas da sua loja**.
5.  **Envio para Diretoria:** Envia um e-mail único para a diretoria contendo:
    * No corpo do e-mail, o destaque de **melhor e pior loja** do dia e do ano (por faturamento).
    * Em anexo, dois rankings em Excel: um de faturamento do dia e outro do ano.
6.  **Backup e Histórico:** Salva as planilhas individuais de cada loja em pastas organizadas por nome da loja e data, criando um histórico de backups.

## Arquivos do Projeto

### Arquivos de Entrada (Input)

Para o script funcionar, ele precisa dos seguintes arquivos (em uma pasta como `/Bases de Dados/`):

* `Emails.xlsx`: Contém nome, loja e e-mail de cada gerente e da diretoria.
* `Vendas.xlsx`: Planilha mestra com o histórico de vendas de **todas** as lojas.
* `Lojas.csv`: Arquivo com o nome de cada loja.

### Arquivos de Saída (Output)

O script gera os seguintes arquivos:

* `/Backup Arquivos Lojas/[Nome da Loja]/[Data] Vendas.xlsx`: Backups diários das vendas por loja.
* `/Rankings/Ranking Dia.xlsx`: Ranking de faturamento do dia.
* `/Rankings/Ranking Ano.xlsx`: Ranking de faturamento do ano.

## Tecnologias Utilizadas

* **Python**
* **Pandas:** Para toda a manipulação, cálculo de indicadores e consolidação dos dados.
* **win32com.client** (ou `smtplib`): Para a automação e envio dos e-mails (com formatação HTML).
* **os** / **pathlib**: Para a manipulação de pastas e arquivos (criação de diretórios de backup).
