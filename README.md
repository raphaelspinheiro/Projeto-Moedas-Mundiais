# 📊 Projeto de Automação de Cotação de Moedas com Python, Selenium, Excel e Power BI

## 📁 Sobre o Projeto

Este projeto tem como objetivo automatizar a coleta diária das cotações das principais moedas estrangeiras, processar e consolidar esses dados em uma planilha Excel, e alimentar um dashboard no Power BI para análise e visualização.

## 🚀 Tecnologias Utilizadas
- `Python 3.13`

- `Selenium` — Para automação da navegação e coleta dos dados na web

- `Pandas` — Para manipulação e tratamento de dados

- `OpenPyXL` — Para edição e atualização de planilhas Excel

- `Excel` — Armazenamento e consolidação dos dados

- `Power BI` — Visualização dos dados em um dashboard interativo

## 🌍 Moedas Monitoradas
- **Dólar Americano (USD)**

- **Euro (EUR)**

- **Libra Esterlina (GBP)**

- **Franco Suíço (CHF)**

- **Iene Japonês (JPY)**

- **Peso Argentino (ARS)**

## 📈 Fluxo do Projeto

**1 - Coleta dos Dados:**
O script em Python usa o Selenium para abrir um navegador e acessar sites confiáveis de câmbio, coletando a cotação atual das moedas listadas.

**2 - Armazenamento em Excel:**
Os dados coletados são armazenados em um arquivo Excel com auxílio do OpenPyXL. Caso o arquivo já exista, os novos dados são adicionados mantendo o histórico.

**3 - Tratamento dos Dados com Pandas:**
Utilizando o Pandas, os dados são limpos, normalizados e organizados em um formato consolidado, prontos para visualização.

**4 - Integração com Power BI:**
A planilha gerada serve como fonte de dados para um dashboard criado no Power BI, onde é possível acompanhar a evolução das cotações de forma visual e dinâmica.

## 📅 Atualizações
Você pode agendar a execução do script com o Agendador de Tarefas do Windows ou cron (Linux) para manter os dados sempre atualizados automaticamente.

---

### 📫 Contato

Caso queira conversar mais sobre o projeto ou colaborar, me chama no [LinkedIn](https://www.linkedin.com/in/raphael-pinheiro-b3062724b)!

---

### 🖥️ Desenvolvido por

Raphael Souza Pinheiro
