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
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
from openpyxl.cell.cell import MergedCell

# Leitura dos dados
notas_df = pd.read_excel("notas.xlsx")
cheques_df = pd.read_excel("cheques.xlsx")

# Padronizar colunas
notas_df.columns = [col.strip().lower().replace(" ", "_") for col in notas_df.columns]
cheques_df.columns = [col.strip().lower().replace(" ", "_") for col in cheques_df.columns]

# Renomear colunas
cheques_df.rename(columns={"valor": "valor_cheque"}, inplace=True)

# Criar lista para armazenar os resultados
resultado = []

cheque_idx = 0
cheque_atual = cheques_df.loc[cheque_idx, "cheque"]
cheque_saldo = cheques_df.loc[cheque_idx, "valor_cheque"]

for _, nota in notas_df.iterrows():
    valor_restante_nf = nota["nf_valor"]
    nf_original = nota["nf"]

    while valor_restante_nf > 0:
        valor_apropriado = min(valor_restante_nf, cheque_saldo)

        resultado.append({
            "fornecedor": nota["fornecedor"],
            "nota_fiscal": nf_original,
            "emissao": nota["emissao"],
            "valor_nf": nota["nf_valor"],
            "valor_apropriado": valor_apropriado,
            "cheque": cheque_atual,
            "valor_cheque": cheques_df.loc[cheque_idx, "valor_cheque"]
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

# Criar DataFrame final e exportar
resultado_df = pd.DataFrame(resultado)
resultado_df.columns = [
    "Fornecedor",
    "Nota fiscal",
    "Emissão",
    "Valor Nota Fiscal",
    "Valor Apropriado",
    "Cheque",
    "Valor Cheque"
]

resultado_df.to_excel("resultado_apropriacao.xlsx", index=False)

# Reabrir com openpyxl para aplicar formatações
wb = load_workbook("resultado_apropriacao.xlsx")
ws = wb.active

# Identificar colunas pelo nome
headers = {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}
col_fornecedor = headers.get("Fornecedor")
col_nf = headers.get("Nota fiscal")
col_emissao = headers.get("Emissão")
col_valor_nf = headers.get("Valor Nota Fiscal")
col_apropriado = headers.get("Valor Apropriado")
col_cheque = headers.get("Cheque")
col_valor_cheque = headers.get("Valor Cheque")

# Estilos
thin_border = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000")
)
center_align = Alignment(horizontal="center", vertical="center")
right_align = Alignment(horizontal="right", vertical="center")
header_fill = PatternFill(start_color="CBCBCB", end_color="CBCBCB", fill_type="solid")
bold_font = Font(bold=True)

# Mesclar células iguais nas colunas 'cheque' e 'valor_cheque'
start_row = 2
prev_cheque = ws.cell(row=start_row, column=col_cheque).value

for row in range(3, ws.max_row + 2):
    current_cheque = ws.cell(row, column=col_cheque).value if row <= ws.max_row else None
    if current_cheque != prev_cheque:
        if row - start_row > 1:
            ws.merge_cells(start_row=start_row, start_column=col_cheque,
                           end_row=row - 1, end_column=col_cheque)
            ws.merge_cells(start_row=start_row, start_column=col_valor_cheque,
                           end_row=row - 1, end_column=col_valor_cheque)
            ws.cell(row=start_row, column=col_cheque).alignment = center_align
            ws.cell(row=start_row, column=col_valor_cheque).alignment = center_align
        start_row = row
        prev_cheque = current_cheque

# Inserir linha de totais
last_data_row = ws.max_row
sum_row = last_data_row + 1

# Mesclar células da coluna A até D para o label 'Total'
ws.merge_cells(start_row=sum_row, start_column=col_fornecedor, end_row=sum_row, end_column=col_valor_nf)
total_cell = ws.cell(row=sum_row, column=col_fornecedor, value="Total")
total_cell.alignment = center_align
total_cell.fill = header_fill
total_cell.font = bold_font
total_cell.border = thin_border

# Inserir fórmulas nas colunas E e G
for col in [col_apropriado, col_valor_cheque]:
    col_letter = get_column_letter(col)
    formula = f"=SUM({col_letter}2:{col_letter}{last_data_row})"
    sum_cell = ws.cell(row=sum_row, column=col, value=formula)
    sum_cell.alignment = center_align
    sum_cell.number_format = '#,##0.00'
    sum_cell.fill = header_fill
    sum_cell.font = bold_font
    sum_cell.border = thin_border

# Aplicar formatação e bordas em todas as células preenchidas
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
    for cell in row:
        if not isinstance(cell, MergedCell) and cell.value is not None:
            cell.border = thin_border

            # Formatar colunas individualmente
            if cell.column == col_nf and isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0'  # Nota fiscal

            elif cell.column == col_emissao:
                cell.number_format = 'DD/MM/YYYY'

            elif cell.column == col_valor_nf and isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00'  # Valor Nota Fiscal

            elif cell.column in [col_apropriado, col_valor_cheque] and isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00'  # Apropriado e Cheque

# Estilizar cabeçalhos
for cell in ws[1]:
    cell.fill = header_fill
    cell.font = bold_font
    cell.alignment = center_align
    cell.border = thin_border

# Salvar
wb.save("resultado_apropriacao.xlsx")
print("✅ Planilha gerada com sucesso com todas as configurações atualizadas!")
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

