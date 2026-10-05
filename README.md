# 📊 Gastos Vini App — Dashboard Poupança Baby

Aplicativo web fullstack (Node.js) para acompanhamento e visualização dos gastos da aba **`Poupança Baby`** (2ª aba visível da planilha Google Sheets), utilizando exclusivamente as **colunas F, G, H, I e J**.

## 🔗 Links de Acesso
- **Produção (VPS Oracle — HTTPS):** [https://poupanca.136-248-111-213.sslip.io](https://poupanca.136-248-111-213.sslip.io)
- **Planilha Base (Google Sheets):** [Abrir Planilha](https://docs.google.com/spreadsheets/d/1Z60EkXO4zn6JEtLMeQ7HvTrIpqeuTgQEt73m2DZX1wY/edit)

## ✨ Funcionalidades
- **Sincronização Ao Vivo com Google Sheets (`Botão Atualizar Planilha`):** O backend Node.js (`server.js`) baixa e processa diretamente a planilha atualizada via `/api/sync` e `/api/expenses`.
- **Mapeamento Estrito das Colunas F a J:**
  - **Coluna F (`situação`):** Status `PAGO` vs `Em aberto`.
  - **Coluna G (`Saídas poupança`):** Valor de cada gasto e soma mensal/geral.
  - **Coluna H (`Dia`) & Coluna I (`Ano`):** Data e ano do gasto (suporta tanto datas seriais quanto períodos descritivos).
  - **Coluna J (`Item`):** Descrição completa do item gasto e categorização automática.
- **Gráfico Interativo de Gastos por Mês:** Clique em qualquer mês no gráfico para filtrar os itens daquele mês.
- **Detalhamento Mensal de Itens:** Lista expansível mês a mês com busca em tempo real, filtro por ano, situação e ordenação por valor ou linha da planilha.

## 🚀 Como Executar Localmente
```bash
node server.js
```
O servidor iniciará em `http://localhost:3075`.
