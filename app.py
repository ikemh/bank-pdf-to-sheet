import pdfplumber
import pandas as pd
import re
import gradio as gr
import tempfile
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

def extrair_movimentacoes_pdf(caminho_pdf):
    registros = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if not texto:
                continue
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
    if not df.empty:
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
                            if cell.value < 0:
                                # Vermelho suave para saídas
                                cell.fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
                            else:
                                # Verde suave para entradas
                                cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

    # Prepara os dados para as colunas independentes de Entrada e Saída
    # "Entrada" = valores > 0; "Saída" = valores negativos (em valor absoluto)
    entradas = [v for v in df["Valor (R$)"] if v > 0]
    saidas = [-v for v in df["Valor (R$)"] if v < 0]

    # Posição inicial para colunas extras (duas colunas à direita da tabela principal, com 1 coluna de espaço)
    extra_start_col = main_table_cols + 2
    col_entrada_letter = get_column_letter(extra_start_col)
    col_saida_letter   = get_column_letter(extra_start_col + 1)
    
    # Estilos para as colunas extras
    extra_header_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    extra_body_fill = PatternFill(start_color="EFF7ED", end_color="EFF7ED", fill_type="solid")
    extra_footer_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    
    # Cabeçalhos das colunas extras
    ws[f"{col_entrada_letter}1"] = "Entrada (R$)"
    ws[f"{col_saida_letter}1"] = "Saída (R$)"
    for col in (f"{col_entrada_letter}1", f"{col_saida_letter}1"):
        ws[col].font = header_font
        ws[col].alignment = center_align
        ws[col].fill = extra_header_fill
        ws[col].border = thin_border

    # Escreve os valores de Entrada e Saída, de forma corrida (sem preencher com zeros)
    max_rows = max(len(entradas), len(saidas))
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

    # Adiciona rodapé (footer) com total de cada coluna extra
    footer_row = max_rows + 2
    ws[f"{col_entrada_letter}{footer_row}"] = "Total"
    ws[f"{col_entrada_letter}{footer_row}"].font = header_font
    ws[f"{col_entrada_letter}{footer_row}"].alignment = center_align
    ws[f"{col_entrada_letter}{footer_row}"].fill = extra_footer_fill
    ws[f"{col_entrada_letter}{footer_row}"].border = thin_border

    ws[f"{col_saida_letter}{footer_row}"] = "Total"
    ws[f"{col_saida_letter}{footer_row}"].font = header_font
    ws[f"{col_saida_letter}{footer_row}"].alignment = center_align
    ws[f"{col_saida_letter}{footer_row}"].fill = extra_footer_fill
    ws[f"{col_saida_letter}{footer_row}"].border = thin_border

    ws[f"{col_entrada_letter}{footer_row + 1}"] = sum(entradas)
    ws[f"{col_entrada_letter}{footer_row + 1}"].number_format = money_format
    ws[f"{col_entrada_letter}{footer_row + 1}"].font = header_font
    ws[f"{col_entrada_letter}{footer_row + 1}"].alignment = center_align
    ws[f"{col_entrada_letter}{footer_row + 1}"].fill = extra_footer_fill
    ws[f"{col_entrada_letter}{footer_row + 1}"].border = thin_border

    ws[f"{col_saida_letter}{footer_row + 1}"] = sum(saidas)
    ws[f"{col_saida_letter}{footer_row + 1}"].number_format = money_format
    ws[f"{col_saida_letter}{footer_row + 1}"].font = header_font
    ws[f"{col_saida_letter}{footer_row + 1}"].alignment = center_align
    ws[f"{col_saida_letter}{footer_row + 1}"].fill = extra_footer_fill
    ws[f"{col_saida_letter}{footer_row + 1}"].border = thin_border

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

def process_pdf(pdf_path):
    if not pdf_path:
        raise gr.Error("Por favor, selecione um arquivo PDF.")
    
    try:
        df = extrair_movimentacoes_pdf(pdf_path)
        
        if df.empty:
            raise ValueError("Não foi possível extrair dados do PDF. Verifique se o formato é suportado.")
            
        # Create a temporary file to save the Excel
        temp_dir = tempfile.gettempdir()
        # Create a user-friendly name for the output file
        base_name = os.path.basename(pdf_path)
        name_without_ext = os.path.splitext(base_name)[0]
        output_filename = f"{name_without_ext}_convertido.xlsx"
        caminho_saida = os.path.join(temp_dir, output_filename)
        
        salvar_excel_formatado(df, caminho_saida)
        
        return caminho_saida
    except Exception as e:
        raise gr.Error(f"Ocorreu um erro: {str(e)}")

# Interface Gradio
with gr.Blocks(title="Conversor de Extrato Bancário (PDF → Excel)") as app:
    gr.Markdown("# 📊 Conversor de Extrato Bancário (PDF → Excel)")
    gr.Markdown("Converta seus extratos bancários em PDF (foco no layout do Sicredi) para uma planilha Excel (.xlsx) formatada, com classificação de entradas/saídas e somatórios.")
    
    with gr.Row():
        with gr.Column():
            pdf_input = gr.File(label="📄 Selecione o PDF do Extrato", file_types=[".pdf"])
            convert_btn = gr.Button("🔄 Converter para Excel", variant="primary")
            
        with gr.Column():
            excel_output = gr.File(label="📥 Arquivo Excel Gerado", interactive=False)
            
    convert_btn.click(
        fn=process_pdf, 
        inputs=[pdf_input], 
        outputs=[excel_output]
    )

if __name__ == "__main__":
    app.launch()
