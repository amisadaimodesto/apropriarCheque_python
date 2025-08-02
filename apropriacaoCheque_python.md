# apropriarCheque_python
# 🧾 Alocador de Cheques por Notas Fiscais (Python)

Este projeto realiza a **apropriação de valores de cheques em notas fiscais** (NFs), respeitando a ordem de entrada e os saldos disponíveis em cada cheque. Ele lê dois arquivos Excel — um com as NFs e outro com os cheques — e gera um terceiro arquivo Excel com os valores apropriados, incluindo mesclagem de células para visualização clara.

---

## 📂 Entrada esperada

- *notas.xlsx* : contendo as colunas `Número NF` e `Valor`
- *cheques.xlsx* : contendo as colunas `Número Cheque` e `Valor`

> As colunas podem conter espaços ou letras maiúsculas — o script padroniza automaticamente.

---

## ⚙️ Intruções

1. Instale as dependências (ou apenas execute o código no compilador de sua escolha (testado no Google Colab com sucesso):
   ```bash
   pip install pandas openpyxl


2. Coloque os arquivos *notas.xlsx* e *cheques.xlsx* na mesma pasta do script (*ou faça o upload dos arquivos no Google Colab*).


3. Execute o script abaixo:
  <pre> \```python # from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# Leitura dos dados com nomes de colunas originais
notas_df = pd.read_excel("notas.xlsx")
cheques_df = pd.read_excel("cheques.xlsx")

# Padronizar colunas para uso interno
notas_df.columns = [col.strip().lower().replace(" ", "_") for col in notas_df.columns]
cheques_df.columns = [col.strip().lower().replace(" ", "_") for col in cheques_df.columns]

# Renomear as colunas para facilitar
notas_df.rename(columns={
    "número_nf": "nf",
    "valor": "valor_nf"
}, inplace=True)

cheques_df.rename(columns={
    "número_cheque": "cheque",
    "valor": "valor_cheque"
}, inplace=True)

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

# Gerar DataFrame com o resultado
resultado_df = pd.DataFrame(resultado)

# Exportar para Excel
output_file = "resultado_apropriacao.xlsx"
resultado_df.to_excel(output_file, index=False)

# MESCLAR células com mesmo cheque consecutivo
wb = load_workbook(output_file)
ws = wb.active

# Encontrar a coluna do "cheque_alocado"
cheque_col_idx = list(resultado_df.columns).index("cheque_alocado") + 1
cheque_col_letter = get_column_letter(cheque_col_idx)

# Mesclar células com mesmo valor consecutivo
start_row = 2
for i in range(3, ws.max_row + 2):  # +2 pois ws.max_row é estático
    current = ws[f"{cheque_col_letter}{i}"].value
    previous = ws[f"{cheque_col_letter}{i - 1}"].value

    if current != previous:
        if i - start_row > 1:
            ws.merge_cells(f"{cheque_col_letter}{start_row}:{cheque_col_letter}{i - 1}")
        start_row = i
else:
    # Último bloco
    if ws.max_row - start_row >= 1:
        ws.merge_cells(f"{cheque_col_letter}{start_row}:{cheque_col_letter}{ws.max_row}")

wb.save(output_file)

print("✅ Apropriação concluída com sucesso! Arquivo salvo como 'resultado_apropriacao.xlsx'") \``` </pre>


4. O arquivo *resultado_apropriacao.xlsx* será gerado com os valores apropriados.


📌 Exemplo de saída
> - A planilha final mostrará:
> - Cada nota fiscal (NF) com seu valor original
> - O valor apropriado de cada cheque
> - Qual cheque foi usado em qual nota
> - Células mescladas para facilitar a leitura dos cheques que cobrem múltiplas NFs\


⚠️ Aviso
Se os cheques não forem suficientes para cobrir todas as NFs, o script exibe um alerta no console:
> - `⚠️ Cheques insuficientes. Algumas NFs podem não ter sido totalmente apropriadas.`


📚 Bibliotecas utilizadas
- `pandas` — para manipulação de dados
- `openpyxl` — para leitura e edição de arquivos Excel

python · pandas · openpyxl · excel · automação · financeiro · notas fiscais · cheques · apropriação
