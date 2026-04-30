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
    saldo_anterior = None
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if not texto:
                continue
            linhas = texto.split('\n')
            for linha in linhas:
                # Procura por SALDO ANTERIOR
                if 'SALDO ANTERIOR' in linha.upper():
                    match_saldo = re.search(r'(-?\d{1,3}(?:\.\d{3})*,\d{2})$', linha)
                    if match_saldo:
                        val_str = match_saldo.group(1).replace('.', '').replace(',', '.')
                        saldo_anterior = float(val_str)
                
                # Filtra somente linhas que iniciem com data
                if not re.match(r'\d{2}/\d{2}/\d{4}', linha):
                    continue
                data = linha[:10]
                restante = linha[11:]
                match = re.search(r'(-?\d{1,3}(?:\.\d{3})*,\d{2})\s+(-?\d{1,3}(?:\.\d{3})*,\d{2})$', restante)
                if match:
                    valor = match.group(1)
                    saldo = match.group(2)
                    descricao = restante[:match.start()].strip()
                    registros.append([data, descricao, valor, saldo])
    df = pd.DataFrame(registros, columns=["Data", "Descrição", "Valor (R$)", "Saldo (R$)"])
    if not df.empty:
        # Converte formato brasileiro para float
        df["Valor (R$)"] = df["Valor (R$)"].str.replace(".", "", regex=False).str.replace(",", ".", regex=False).astype(float)
        df["Saldo (R$)"] = df["Saldo (R$)"].str.replace(".", "", regex=False).str.replace(",", ".", regex=False).astype(float)
    return df, saldo_anterior

def salvar_excel_formatado(df, caminho_saida, saldo_anterior=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Extrato"

    # Escrita da tabela principal (A1 em diante)
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    
    # Adiciona a coluna E (Categoria)
    ws["E1"] = "Categoria"
    
    main_table_end = ws.max_row
    
    # Preenche fórmulas da Categoria
    for row_idx in range(2, main_table_end + 1):
        ws[f"E{row_idx}"] = f'=IF(UPPER(TRIM(B{row_idx}))="RESG.APLIC.FIN.AVISO PREV CAPTACAO","Resgate",IF(C{row_idx}>0,"Entrada",IF(C{row_idx}<0,"Saída","")))'

    main_table_cols = 5 # A, B, C, D, E

    # Estilos básicos
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
    for row in ws.iter_rows(min_row=1, max_row=main_table_end, max_col=main_table_cols):
        for cell in row:
            cell.alignment = center_align
            cell.border = thin_border
            if cell.row == 1:
                cell.font = header_font
                cell.fill = main_header_fill
            else:
                cell.fill = main_body_fill
                if cell.column_letter in ['C', 'D']:
                    cell.number_format = money_format
                
                # Cores dinâmicas na coluna C baseadas na descrição
                if cell.column_letter == 'C':
                    descricao_val = ws.cell(row=cell.row, column=2).value
                    is_resgate = descricao_val and str(descricao_val).strip().upper() == "RESG.APLIC.FIN.AVISO PREV CAPTACAO"
                    if is_resgate:
                        cell.fill = main_body_fill
                    elif cell.value is not None and cell.value < 0:
                        cell.fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
                    elif cell.value is not None and cell.value > 0:
                        cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

    # Congelar cabeçalho e AutoFilter
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:E{main_table_end}"

    # Posição inicial para colunas extras (1 coluna em branco)
    extra_start_col = main_table_cols + 2
    col_entrada_letter = get_column_letter(extra_start_col)
    col_saida_letter   = get_column_letter(extra_start_col + 1)
    col_resgate_letter = get_column_letter(extra_start_col + 2)
    
    # Estilos extras
    extra_header_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    extra_body_fill = PatternFill(start_color="EFF7ED", end_color="EFF7ED", fill_type="solid")
    extra_footer_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    
    # Cabeçalhos extras
    ws[f"{col_entrada_letter}1"] = "Entrada (R$)"
    ws[f"{col_saida_letter}1"] = "Saída (R$)"
    ws[f"{col_resgate_letter}1"] = "Resgates (R$)"
    
    for col_letter in (col_entrada_letter, col_saida_letter, col_resgate_letter):
        ws[f"{col_letter}1"].font = header_font
        ws[f"{col_letter}1"].alignment = center_align
        ws[f"{col_letter}1"].fill = extra_header_fill
        ws[f"{col_letter}1"].border = thin_border

    # Fórmulas de Entrada, Saída e Resgate linha por linha
    for row_idx in range(2, main_table_end + 1):
        ws[f"{col_entrada_letter}{row_idx}"] = f'=IF($E{row_idx}="Entrada",$C{row_idx},"")'
        ws[f"{col_saida_letter}{row_idx}"] = f'=IF($E{row_idx}="Saída",ABS($C{row_idx}),"")'
        ws[f"{col_resgate_letter}{row_idx}"] = f'=IF($E{row_idx}="Resgate",$C{row_idx},"")'
        
        for col_letter in (col_entrada_letter, col_saida_letter, col_resgate_letter):
            cell = ws[f"{col_letter}{row_idx}"]
            cell.number_format = money_format
            cell.alignment = center_align
            cell.fill = extra_body_fill
            cell.border = thin_border

    # Totais (SUM)
    total_row = main_table_end + 2
    for col_letter in (col_entrada_letter, col_saida_letter, col_resgate_letter):
        ws[f"{col_letter}{total_row - 1}"] = "Total"
        ws[f"{col_letter}{total_row - 1}"].font = header_font
        ws[f"{col_letter}{total_row - 1}"].alignment = center_align
        ws[f"{col_letter}{total_row - 1}"].fill = extra_footer_fill
        ws[f"{col_letter}{total_row - 1}"].border = thin_border

        ws[f"{col_letter}{total_row}"] = f"=SUM({col_letter}2:{col_letter}{main_table_end})"
        ws[f"{col_letter}{total_row}"].number_format = money_format
        ws[f"{col_letter}{total_row}"].font = header_font
        ws[f"{col_letter}{total_row}"].alignment = center_align
        ws[f"{col_letter}{total_row}"].fill = extra_footer_fill
        ws[f"{col_letter}{total_row}"].border = thin_border

    # Seção de Conferência do Balancete
    conf_start_col = extra_start_col + 4 # Deixa uma coluna em branco
    col_label_letter = get_column_letter(conf_start_col)
    col_val_letter = get_column_letter(conf_start_col + 1)
    
    conf_row = 2
    # Título
    ws[f"{col_label_letter}{conf_row}"] = "Conferência do Balancete"
    ws[f"{col_label_letter}{conf_row}"].font = header_font
    ws.merge_cells(f"{col_label_letter}{conf_row}:{col_val_letter}{conf_row}")
    conf_row += 1

    campos = [
        "Saldo anterior",
        "Saldo final",
        "Variação do saldo",
        "Total entradas",
        "Total saídas",
        "Total resgates",
        "Resultado líquido",
        "Diferença de conferência",
        "Status"
    ]
    
    # Prepara as células da conferência
    conf_cells = {}
    for campo in campos:
        ws[f"{col_label_letter}{conf_row}"] = campo
        ws[f"{col_label_letter}{conf_row}"].font = Font(bold=True)
        conf_cells[campo] = f"{col_val_letter}{conf_row}"
        conf_row += 1
        
    # Fórmulas da Conferência
    if saldo_anterior is not None:
        ws[conf_cells["Saldo anterior"]] = saldo_anterior
    
    ws[conf_cells["Saldo final"]] = f"=D{main_table_end}"
    ws[conf_cells["Variação do saldo"]] = f"={conf_cells['Saldo final']}-{conf_cells['Saldo anterior']}"
    ws[conf_cells["Total entradas"]] = f"={col_entrada_letter}{total_row}"
    ws[conf_cells["Total saídas"]] = f"={col_saida_letter}{total_row}"
    ws[conf_cells["Total resgates"]] = f"={col_resgate_letter}{total_row}"
    ws[conf_cells["Resultado líquido"]] = f"={conf_cells['Total entradas']}+{conf_cells['Total resgates']}-{conf_cells['Total saídas']}"
    ws[conf_cells["Diferença de conferência"]] = f"={conf_cells['Resultado líquido']}-{conf_cells['Variação do saldo']}"
    ws[conf_cells["Status"]] = f'=IF(ABS({conf_cells["Diferença de conferência"]})<0.01,"OK","DIVERGENTE")'

    # Formatação das células de conferência
    for campo, cel in conf_cells.items():
        cell = ws[cel]
        cell.alignment = Alignment(horizontal="right")
        if campo != "Status":
            cell.number_format = money_format
        else:
            cell.font = Font(bold=True)

    # Ajuste automático da largura das colunas
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            if cell.value is not None:
                # Evita erro em fórmulas longas ao medir largura
                val_str = str(cell.value)
                if val_str.startswith('='):
                    val_str = "R$ 99.999,99"
                max_length = max(max_length, len(val_str))
        adjusted_width = max_length + 6
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
        df, saldo_anterior = extrair_movimentacoes_pdf(entrada)
        salvar_excel_formatado(df, saida, saldo_anterior)
        messagebox.showinfo("Sucesso", "Arquivo convertido e salvo com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

if __name__ == "__main__":
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
