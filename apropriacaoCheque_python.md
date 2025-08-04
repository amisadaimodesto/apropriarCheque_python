# apropriarCheque_python
# 🧾 Alocador de Cheques por Notas Fiscais #
- ![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
- ![Pandas](https://img.shields.io/badge/Pandas-150458?logo=pandas&logoColor=white)

Este projeto realiza a **apropriação de valores de cheques em notas fiscais** (NFs), respeitando a ordem de entrada e os saldos disponíveis em cada cheque. Ele lê dois arquivos Excel — um com as NFs e outro com os cheques — e gera um terceiro arquivo Excel com os valores apropriados, incluindo mesclagem de células para visualização clara.

---

## 📂 Entrada esperada

- *notas.xlsx* : contendo as colunas `Número NF` e `Valor`
- *cheques.xlsx* : contendo as colunas `Número Cheque` e `Valor`

> As colunas podem conter espaços ou letras maiúsculas — o script padroniza automaticamente.

---

## ⚙️ Intruções

1. Instale as dependências (ou apenas execute o código no compilador de sua escolha (*testado no Google Colab com sucesso*):
   ```bash
   pip install pandas openpyxl


2. Coloque os arquivos *notas.xlsx* e *cheques.xlsx* na mesma pasta do script (*ou faça o upload dos arquivos no Google Colab*).


3. Execute o script abaixo:
```python
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, numbers
from openpyxl.utils import get_column_letter

# Leitura dos dados
notas_df = pd.read_excel("notas.xlsx")
cheques_df = pd.read_excel("cheques.xlsx")

# Padronizar colunas
notas_df.columns = [col.strip().lower().replace(" ", "_") for col in notas_df.columns]
cheques_df.columns = [col.strip().lower().replace(" ", "_") for col in cheques_df.columns]

# Renomear colunas se necessário
cheques_df.rename(columns={"valor": "valor_cheque"}, inplace=True)

# Criar lista para armazenar as linhas resultantes
resultado = []

cheque_idx = 0
cheque_atual = cheques_df.loc[cheque_idx, "cheque"]
cheque_saldo = cheques_df.loc[cheque_idx, "valor_cheque"]

for _, nota in notas_df.iterrows():
    valor_restante_nf = nota["valor_nf"]
    nf_original = nota["nf"]

    while valor_restante_nf > 0:
        valor_apropriado = min(valor_restante_nf, cheque_saldo)

        resultado.append({
            "nf": nf_original,
            "valor_nf": nota["valor_nf"],
            "valor_apropriado": valor_apropriado,
            "cheque_alocado": cheque_atual,
            "valor_do_cheque": cheques_df.loc[cheque_idx, "valor_cheque"]
        })

        valor_restante_nf -= valor_apropriado
        cheque_saldo -= valor_apropriado

        if cheque_saldo == 0:
            cheque_idx += 1
            if cheque_idx >= len(cheques_df):
                print("⚠️ Cheques insuficientes. Algumas NFs podem não ter sido totalmente apropriadas.")
                break
            cheque_atual = cheques_df.loc[cheque_idx, "cheque"]
            cheque_saldo = cheques_df.loc[cheque_idx, "valor_cheque"]

    if cheque_idx >= len(cheques_df):
        break

# Criar DataFrame e exportar
resultado_df = pd.DataFrame(resultado)
resultado_df.to_excel("resultado_apropriacao.xlsx", index=False)

# Reabrir com openpyxl para aplicar formatação e mesclagens
wb = load_workbook("resultado_apropriacao.xlsx")
ws = wb.active

# Localizar os índices das colunas
headers = {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}
col_cheque = headers.get("cheque_alocado")
col_valor = headers.get("valor_do_cheque")
col_valor_aprop = headers.get("valor_apropriado")

# Formatando as colunas com valores monetários (com separador de milhar, vírgula como decimal)
valor_format = '#,##0.00'  # Excel aplica ponto como separador de milhar, mas isso será formatado conforme local do Excel

for row in range(2, ws.max_row + 1):
    if col_valor_aprop:
        cell_aprop = ws.cell(row=row, column=col_valor_aprop)
        cell_aprop.number_format = valor_format
    if col_valor:
        cell_valor = ws.cell(row=row, column=col_valor)
        cell_valor.number_format = valor_format

# Mesclar células para cheques e valores correspondentes
if col_cheque and col_valor:
    start_row = 2
    prev_cheque = ws.cell(row=start_row, column=col_cheque).value

    for row in range(3, ws.max_row + 2):  # +2 para forçar o último grupo
        current_cheque = ws.cell(row, column=col_cheque).value if row <= ws.max_row else None
        if current_cheque != prev_cheque:
            if row - start_row > 1:
                # Mesclar colunas cheque_alocado e valor_do_cheque
                ws.merge_cells(start_row=start_row, start_column=col_cheque,
                               end_row=row - 1, end_column=col_cheque)
                ws.merge_cells(start_row=start_row, start_column=col_valor,
                               end_row=row - 1, end_column=col_valor)

                # Centralizar conteúdo das células mescladas
                merged_cell_cheque = ws.cell(row=start_row, column=col_cheque)
                merged_cell_cheque.alignment = Alignment(horizontal='center', vertical='center')

                merged_cell_valor = ws.cell(row=start_row, column=col_valor)
                merged_cell_valor.alignment = Alignment(horizontal='center', vertical='center')

            start_row = row
            prev_cheque = current_cheque

wb.save("resultado_apropriacao.xlsx")
print("✅ Planilha formatada com sucesso e salva como 'resultado_apropriacao.xlsx'")
```

4. O arquivo *resultado_apropriacao.xlsx* será gerado com os valores apropriados.


📌 Exemplo de saída
> - A planilha final mostrará:
> - Cada nota fiscal (NF) com seu valor original;
> - O valor apropriado de cada cheque;
> - Qual cheque foi usado em qual nota;
> - Células mescladas para facilitar a leitura dos cheques que cobrem múltiplas NFs.


⚠️ Aviso
Se os cheques não forem suficientes para cobrir todas as NFs, o script exibe um alerta no console:
> - `⚠️ Cheques insuficientes. Algumas NFs podem não ter sido totalmente apropriadas.`


📚 Bibliotecas utilizadas
- `pandas` — para manipulação de dados
- `openpyxl` — para leitura e edição de arquivos Excel

#python · #pandas · #openpyxl · #excel · #automação · #financeiro · #notas fiscais · #cheques · #apropriação

