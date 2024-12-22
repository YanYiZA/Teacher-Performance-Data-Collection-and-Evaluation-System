# -*- coding: utf-8 -*-
import sys
import os
from openpyxl import Workbook
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ���ڴ���Excel����
def export_to_excel(output_file):
    # ����һ���µ�Excel�������͹�����
    wb = Workbook()
    ws = wb.active

    # ��data.txt�ļ�����ȡ����
    with open('data.txt', 'r', encoding='utf-8') as f:  # ����data.txt�ļ��ڵ�ǰĿ¼
        lines = f.readlines()

    # ��txt�ļ�������д��Excel
    for row_num, line in enumerate(lines, start=1):
        cells = line.strip().split('\t')  # ���Ʊ���ָ�ÿһ��
        for col_num, cell in enumerate(cells, start=1):
            ws.cell(row=row_num, column=col_num, value=cell)  # д�뵥Ԫ��

    # ���浽ָ����Excel�ļ�
    wb.save(output_file)
    print(f"Excel file saved: {output_file}")

# ���ڴ���Word����
def export_to_word(output_file):
    # ����һ���µ�Word�ĵ�
    doc = Document()

    # Ԥ�����ͷ
    columns = ["Name", "Teaching Work", "Research Work", "Scientific Research", "Other Work", "Total Performance"]

    # ��data.txt�ļ�����ȡ����
    with open('data.txt', 'r', encoding='utf-8') as f:  # ����data.txt�ļ��ڵ�ǰĿ¼
        lines = f.readlines()

    # ����һ����񣬼�������û���ر��ӵĽṹ
    table = doc.add_table(rows=1, cols=len(columns))

    # �����ı�ͷ
    hdr_cells = table.rows[0].cells
    for i, header in enumerate(columns):
        hdr_cells[i].text = header

    # ���������
    for line in lines:
        row_cells = table.add_row().cells
        cells = line.strip().split('\t')  # ���Ʊ���ָ�ÿһ��
        for i, cell in enumerate(cells):
            row_cells[i].text = cell

    # ���浽ָ����Word�ļ�
    doc.save(output_file)
    print(f"Word file saved: {output_file}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python export_data.py <format> <output_file>")
        sys.exit(1)

    # ��ȡ�����в���
    file_format = sys.argv[1].lower()  # 'excel' �� 'word'
    output_file = sys.argv[2]

    # ȷ��txt�ļ�����
    if not os.path.exists('data.txt'):
        print(f"Error: The file data.txt does not exist.")
        sys.exit(1)

    # ���ݸ�ʽѡ����ʽ
    if file_format == 'excel':
        export_to_excel(output_file)
    elif file_format == 'word':
        export_to_word(output_file)
    else:
        print("Error: Unsupported format. Please specify 'excel' or 'word'.")
        sys.exit(1)
