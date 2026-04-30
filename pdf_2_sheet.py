import pdfplumber
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

def extrair_movimentacoes_pdf(caminho_pdf):
    registros = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            linhas = texto.split('\n')
            for linha in linhas:
                # Filtra somente linhas que iniciem com data
                if not re.match(r'\d{2}/\d{2}/\d{4}', linha):
                    continue
                data = linha[:10]
                restante = linha[11:]
                match = re.search(r'(-?\d{1,3}(?:\.\d{3})*,\d{2})\s+(\d{1,3}(?:\.\d{3})*,\d{2})$', restante)
                if match:
                    valor = match.group(1)
                    saldo = match.group(2)
                    descricao = restante[:match.start()].strip()
                    registros.append([data, descricao, valor, saldo])
    df = pd.DataFrame(registros, columns=["Data", "Descrição", "Valor (R$)", "Saldo (R$)"])
    # Converte formato brasileiro para float
    df["Valor (R$)"] = df["Valor (R$)"].str.replace(".", "", regex=False).str.replace(",", ".", regex=False).astype(float)
    df["Saldo (R$)"] = df["Saldo (R$)"].str.replace(".", "", regex=False).str.replace(",", ".", regex=False).astype(float)
    return df

def salvar_excel_formatado(df, caminho_saida):
    # Cria workbook novo
    wb = Workbook()
    ws = wb.active
    ws.title = "Extrato"

    # Escrita da tabela principal (A1 em diante)
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    
    # Estilos para a tabela principal
    main_header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    main_body_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    header_font = Font(bold=True)
    center_align = Alignment(horizontal="center")
    thin_border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC'),
    )
    money_format = u'R$ #,##0.00'

    # Aplica estilos na tabela principal
    main_table_end = ws.max_row
    main_table_cols = ws.max_column

    for row in ws.iter_rows(min_row=1, max_row=main_table_end, max_col=main_table_cols):
        for cell in row:
            cell.alignment = center_align
            cell.border = thin_border
            if cell.row == 1:  # Cabeçalho
                cell.font = header_font
                cell.fill = main_header_fill
            else:
                cell.fill = main_body_fill
                if isinstance(cell.value, float):
                    cell.number_format = money_format

                    if cell.column_letter == 'C':
                            descricao_val = ws.cell(row=cell.row, column=2).value
                            is_resgate = descricao_val and str(descricao_val).strip().upper() == "RESG.APLIC.FIN.AVISO PREV CAPTACAO"
                            
                            if is_resgate:
                                # Neutro para resgates
                                cell.fill = main_body_fill
                            elif cell.value < 0:
                                # Vermelho suave para saídas
                                cell.fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
                            else:
                                # Verde suave para entradas
                                cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

    # Prepara os dados para as colunas independentes de Entrada, Saída e Resgates
    DESCRICAO_RESGATE = "RESG.APLIC.FIN.AVISO PREV CAPTACAO"
    
    resgates = [
        row["Valor (R$)"]
        for _, row in df.iterrows()
        if row["Descrição"].strip().upper() == DESCRICAO_RESGATE
    ]

    entradas = [
        row["Valor (R$)"]
        for _, row in df.iterrows()
        if row["Valor (R$)"] > 0
        and row["Descrição"].strip().upper() != DESCRICAO_RESGATE
    ]

    saidas = [
        -row["Valor (R$)"]
        for _, row in df.iterrows()
        if row["Valor (R$)"] < 0
        and row["Descrição"].strip().upper() != DESCRICAO_RESGATE
    ]

    # Posição inicial para colunas extras (duas colunas à direita da tabela principal, com 1 coluna de espaço)
    extra_start_col = main_table_cols + 2
    col_entrada_letter = get_column_letter(extra_start_col)
    col_saida_letter   = get_column_letter(extra_start_col + 1)
    col_resgate_letter = get_column_letter(extra_start_col + 2)
    
    # Estilos para as colunas extras
    extra_header_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    extra_body_fill = PatternFill(start_color="EFF7ED", end_color="EFF7ED", fill_type="solid")
    extra_footer_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    
    # Cabeçalhos das colunas extras
    ws[f"{col_entrada_letter}1"] = "Entrada (R$)"
    ws[f"{col_saida_letter}1"] = "Saída (R$)"
    ws[f"{col_resgate_letter}1"] = "Resgates (R$)"
    for col in (f"{col_entrada_letter}1", f"{col_saida_letter}1", f"{col_resgate_letter}1"):
        ws[col].font = header_font
        ws[col].alignment = center_align
        ws[col].fill = extra_header_fill
        ws[col].border = thin_border

    # Escreve os valores de Entrada, Saída e Resgates, de forma corrida (sem preencher com zeros)
    max_rows = max(len(entradas), len(saidas), len(resgates))
    for i in range(max_rows):
        row_idx = i + 2  # Inicia na linha 2
        if i < len(entradas):
            cell_entrada = ws[f"{col_entrada_letter}{row_idx}"]
            cell_entrada.value = entradas[i]
            cell_entrada.number_format = money_format
            cell_entrada.alignment = center_align
            cell_entrada.fill = extra_body_fill
            cell_entrada.border = thin_border
        if i < len(saidas):
            cell_saida = ws[f"{col_saida_letter}{row_idx}"]
            cell_saida.value = saidas[i]
            cell_saida.number_format = money_format
            cell_saida.alignment = center_align
            cell_saida.fill = extra_body_fill
            cell_saida.border = thin_border
        if i < len(resgates):
            cell_resgate = ws[f"{col_resgate_letter}{row_idx}"]
            cell_resgate.value = resgates[i]
            cell_resgate.number_format = money_format
            cell_resgate.alignment = center_align
            cell_resgate.fill = extra_body_fill
            cell_resgate.border = thin_border

    # Adiciona rodapé (footer) com total de cada coluna extra
    footer_row = max_rows + 2
    for col_letter, total_val in [
        (col_entrada_letter, sum(entradas)), 
        (col_saida_letter, sum(saidas)), 
        (col_resgate_letter, sum(resgates))
    ]:
        ws[f"{col_letter}{footer_row}"] = "Total"
        ws[f"{col_letter}{footer_row}"].font = header_font
        ws[f"{col_letter}{footer_row}"].alignment = center_align
        ws[f"{col_letter}{footer_row}"].fill = extra_footer_fill
        ws[f"{col_letter}{footer_row}"].border = thin_border

        ws[f"{col_letter}{footer_row + 1}"] = total_val
        ws[f"{col_letter}{footer_row + 1}"].number_format = money_format
        ws[f"{col_letter}{footer_row + 1}"].font = header_font
        ws[f"{col_letter}{footer_row + 1}"].alignment = center_align
        ws[f"{col_letter}{footer_row + 1}"].fill = extra_footer_fill
        ws[f"{col_letter}{footer_row + 1}"].border = thin_border

    # Ajuste automático da largura das colunas para acomodar o conteúdo
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Identifica a letra da coluna
        for cell in col:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))
        adjusted_width = max_length + 10  # Adiciona um buffer
        ws.column_dimensions[column].width = adjusted_width

    wb.save(caminho_saida)

def selecionar_pdf():
    caminho = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    entrada_var.set(caminho)

def selecionar_saida():
    caminho = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    saida_var.set(caminho)

def converter():
    entrada = entrada_var.get()
    saida = saida_var.get()
    if not entrada or not saida:
        messagebox.showerror("Erro", "Por favor, selecione o PDF de entrada e o caminho de saída.")
        return
    try:
        df = extrair_movimentacoes_pdf(entrada)
        salvar_excel_formatado(df, saida)
        messagebox.showinfo("Sucesso", "Arquivo convertido e salvo com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

# Interface gráfica (GUI)
janela = tk.Tk()
janela.title("Conversor de Extrato Bancário (PDF → Excel)")

entrada_var = tk.StringVar()
saida_var = tk.StringVar()

tk.Label(janela, text="Selecionar PDF:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
tk.Entry(janela, textvariable=entrada_var, width=50).grid(row=0, column=1, padx=5)
tk.Button(janela, text="Buscar", command=selecionar_pdf).grid(row=0, column=2, padx=5)

tk.Label(janela, text="Salvar como:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
tk.Entry(janela, textvariable=saida_var, width=50).grid(row=1, column=1, padx=5)
tk.Button(janela, text="Escolher", command=selecionar_saida).grid(row=1, column=2, padx=5)

tk.Button(janela, text="Converter", command=converter, bg="#4CAF50", fg="white", height=2).grid(row=2, column=1, pady=20)

janela.mainloop()
