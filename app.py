import streamlit as st
import zipfile
import os
from PyPDF2 import PdfReader
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import re

def extrair_dados_pdf(file):
    reader = PdfReader(file)
    full_text = ""
    for page in reader.pages:
        text = page.extract_text()
        if text:
            full_text += "\n" + text

    dados = {
        "NumeroNota": None,
        "DataEmissao": None,
        "CNPJ": None,
        "NomeTitular": None,
        "ValorTotal": None
    }

    match_nota = re.search(r'NOTA\s*FISCAL\s*[N\u00b0NO]*\s*(\d+)', full_text, re.IGNORECASE)
    if match_nota:
        dados["NumeroNota"] = match_nota.group(1)

    match_data = re.search(r'DATA\s*DE\s*EMISS\u00c3O\s*[:\s]*(\d{2}/\d{2}/\d{4})', full_text, re.IGNORECASE)
    if match_data:
        dados["DataEmissao"] = match_data.group(1)

    linhas = full_text.split('\n')
    for idx, linha in enumerate(linhas):
        if "CNPJ/CPF:" in linha:
            match_cnpj = re.search(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', linha)
            if match_cnpj:
                dados["CNPJ"] = match_cnpj.group(0)
            if idx > 0:
                dados["NomeTitular"] = linhas[idx-1].strip()
            break

    valores = re.findall(r'R\$\**\s*\d{1,3}(?:\.\d{3})*,\d{2}', full_text)
    if valores:
        valores_float = []
        for valor in valores:
            valor_limpo = re.sub(r'[^\d,]', '', valor)
            valor_float = float(valor_limpo.replace('.', '').replace(',', '.'))
            valores_float.append(valor_float)
        if valores_float:
            dados["ValorTotal"] = max(valores_float)

    return dados

def gerar_planilha(dados_extraidos):
    wb = Workbook()
    wb.remove(wb.active)
    headers = ["Número da Nota", "Data de Emissão", "CNPJ do Titular", "Nome do Titular", "Valor Total NF"]

    for uc, dados in dados_extraidos.items():
        ws = wb.create_sheet(title=uc)
        ws.append(headers)

        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4F4F4F")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        valores = [
            dados.get("NumeroNota"),
            dados.get("DataEmissao"),
            dados.get("CNPJ"),
            dados.get("NomeTitular"),
            dados.get("ValorTotal")
        ]

        for col_num, valor in enumerate(valores, start=1):
            ws.cell(row=2, column=col_num, value=valor)
            ws.cell(row=2, column=col_num).alignment = Alignment(horizontal="center", vertical="center")

        for col in ws.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
            adjusted_width = (max_length + 4)
            ws.column_dimensions[col_letter].width = adjusted_width

    return wb

def main():
    st.title("💡 Extrator de Dados - Notas Fiscais Energia Elétrica")
    uploaded_files = st.file_uploader("Envie os PDFs aqui:", type="pdf", accept_multiple_files=True)

    if uploaded_files:
        dados_extraidos = {}

        for uploaded_file in uploaded_files:
            unidade_consumidora = re.search(r'(\d{6,})', uploaded_file.name)
            if unidade_consumidora:
                uc_number = unidade_consumidora.group(1)
                dados = extrair_dados_pdf(uploaded_file)
                dados_extraidos[uc_number] = dados

        wb = gerar_planilha(dados_extraidos)
        output_path = "/tmp/Planilha_Resultante.xlsx"
        wb.save(output_path)

        with open(output_path, "rb") as file:
            st.download_button(
                label="📂 Baixar Planilha Excel",
                data=file,
                file_name="Planilha_Resultante.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()