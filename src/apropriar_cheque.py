from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
from openpyxl.cell.cell import MergedCell

BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = BASE_DIR / "data" / "input"
OUTPUT_DIR = BASE_DIR / "data" / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Leitura dos dados
notas_df = pd.read_excel(INPUT_DIR / "notas.xlsx")
cheques_df = pd.read_excel(INPUT_DIR / "cheques.xlsx")

# Padronizar colunas
notas_df.columns = [col.strip().lower().replace(" ", "_") for col in notas_df.columns]
cheques_df.columns = [col.strip().lower().replace(" ", "_") for col in cheques_df.columns]
cheques_df.rename(columns={"valor": "valor_cheque"}, inplace=True)

# Motor de apropriação FIFO
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

output_file = OUTPUT_DIR / "resultado_apropriacao.xlsx"
resultado_df.to_excel(output_file, index=False)

# Formatação contábil com OpenPyXL
wb = load_workbook(output_file)
ws = wb.active

headers = {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}
col_fornecedor = headers.get("Fornecedor")
col_nf = headers.get("Nota fiscal")
col_emissao = headers.get("Emissão")
col_valor_nf = headers.get("Valor Nota Fiscal")
col_apropriado = headers.get("Valor Apropriado")
col_cheque = headers.get("Cheque")
col_valor_cheque = headers.get("Valor Cheque")

thin_border = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000")
)
center_align = Alignment(horizontal="center", vertical="center")
header_fill = PatternFill(start_color="CBCBCB", end_color="CBCBCB", fill_type="solid")
bold_font = Font(bold=True)

start_row = 2
prev_cheque = ws.cell(row=start_row, column=col_cheque).value

for row in range(3, ws.max_row + 2):
    current_cheque = ws.cell(row, column=col_cheque).value if row <= ws.max_row else None
    if current_cheque != prev_cheque:
        if row - start_row > 1:
            ws.merge_cells(start_row=start_row, start_column=col_cheque, end_row=row - 1, end_column=col_cheque)
            ws.merge_cells(start_row=start_row, start_column=col_valor_cheque, end_row=row - 1, end_column=col_valor_cheque)
            ws.cell(row=start_row, column=col_cheque).alignment = center_align
            ws.cell(row=start_row, column=col_valor_cheque).alignment = center_align
        start_row = row
        prev_cheque = current_cheque

last_data_row = ws.max_row
sum_row = last_data_row + 1

ws.merge_cells(start_row=sum_row, start_column=col_fornecedor, end_row=sum_row, end_column=col_valor_nf)
total_cell = ws.cell(row=sum_row, column=col_fornecedor, value="Total")
total_cell.alignment = center_align
total_cell.fill = header_fill
total_cell.font = bold_font
total_cell.border = thin_border

for col in [col_apropriado, col_valor_cheque]:
    col_letter = get_column_letter(col)
    formula = f"=SUM({col_letter}2:{col_letter}{last_data_row})"
    sum_cell = ws.cell(row=sum_row, column=col, value=formula)
    sum_cell.alignment = center_align
    sum_cell.number_format = '#,##0.00'
    sum_cell.fill = header_fill
    sum_cell.font = bold_font
    sum_cell.border = thin_border

for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
    for cell in row:
        if not isinstance(cell, MergedCell) and cell.value is not None:
            cell.border = thin_border
            if cell.column == col_nf and isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0'
            elif cell.column == col_emissao:
                cell.number_format = 'DD/MM/YYYY'
            elif cell.column in [col_valor_nf, col_apropriado, col_valor_cheque] and isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00'

for cell in ws[1]:
    cell.fill = header_fill
    cell.font = bold_font
    cell.alignment = center_align
    cell.border = thin_border

wb.save(output_file)
print(f"✅ Planilha gerada com sucesso em: {output_file}")
