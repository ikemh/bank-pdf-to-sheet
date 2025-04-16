# Bank PDF to Sheet

Conversor de extratos bancários em PDF para planilhas Excel (.xlsx), com foco atual no layout dos extratos do banco Sicredi.

## 📄 Descrição

Este utilitário permite a extração de dados tabulares a partir de PDFs de extrato bancário e converte em uma planilha com formatação legível, incluindo classificação de entradas/saídas e somatórios.

> Atualmente compatível apenas com o layout textual dos extratos do **Sicredi**.  
> Outros bancos podem ser suportados com ajustes no parser de texto.

---

## ✅ Funcionalidades

- Interface gráfica (Tkinter) para seleção de arquivo PDF e caminho de saída
- Conversão precisa dos valores (R$) com detecção de entradas e saídas
- Geração de colunas auxiliares com somatórios
- Estilização visual para facilitar leitura
- Compatível com Windows (build via PyInstaller)

---

## 🚀 Requisitos

- Python 3.11+
- Dependências:
  - `pdfplumber`
  - `pandas`
  - `openpyxl`

Instale com:

```bash
pip install -r requirements.txt
```

---

## 🖥️ Execução

```bash
python pdf_2_sheet.py
```

---

## 📦 Geração de Executável

Para empacotar com PyInstaller:

```bash
pyinstaller --clean --windowed --icon=icone.ico ^
  --distpath dist pdf_2_sheet.py
```

O executável será gerado na pasta `dist/`.

---

## 📁 Estrutura

```
bank-pdf-to-sheet/
├── pdf_2_sheet.py            # Script principal com GUI e processamento
├── pdf_2_sheet.spec          # Configuração do PyInstaller
├── dist/                     # Executável gerado
├── build/                    # Arquivos temporários da build
└── requirements.txt          # Lista de dependências
```

---

## 📌 Observações

- O PDF precisa ser **baseado em texto extraível** (não digitalizado por imagem).
- Testado com extratos do Sicredi no layout exportado via internet banking.

---

## 📄 Licença

Uso pessoal ou corporativo privado.  
Não autorizado para redistribuição pública sem consentimento do autor.