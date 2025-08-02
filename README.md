# apropriarCheque_python
# 🧾 Alocador de Cheques por Notas Fiscais (Python)

Este projeto realiza a **apropriação de valores de cheques em notas fiscais** (NFs), respeitando a ordem de entrada e os saldos disponíveis em cada cheque. Ele lê dois arquivos Excel — um com as NFs e outro com os cheques — e gera um terceiro arquivo Excel com os valores apropriados, incluindo mesclagem de células para visualização clara.

---

## 📂 Entrada esperada

- * notas.xlsx * : contendo as colunas `Número NF` e `Valor`
- * cheques.xlsx * : contendo as colunas `Número Cheque` e `Valor`

> As colunas podem conter espaços ou letras maiúsculas — o script padroniza automaticamente.

---

## ⚙️ Como usar

1. Instale as dependências:
   ```bash
   pip install pandas openpyxl

2. Coloque os arquivos **notas.xlsx** e **cheques.xlsx** na mesma pasta do script.

3. Execute o script:
  ** python apropriacao.py **

4. O arquivo ** resultado_apropriacao.xlsx ** será gerado com os valores apropriados.

📌 Exemplo de saída
> - A planilha final mostrará:
> - Cada nota fiscal (NF) com seu valor original
> - O valor apropriado de cada cheque
> - Qual cheque foi usado em qual nota
> - Células mescladas para facilitar a leitura dos cheques que cobrem múltiplas NFs

⚠️ Aviso
Se os cheques não forem suficientes para cobrir todas as NFs, o script exibe um alerta no console:
> - `⚠️ Cheques insuficientes. Algumas NFs podem não ter sido totalmente apropriadas.`

📚 Bibliotecas utilizadas
- `pandas` — para manipulação de dados
- `openpyxl` — para leitura e edição de arquivos Excel

python · pandas · openpyxl · excel · automação · financeiro · notas fiscais · cheques · apropriação
