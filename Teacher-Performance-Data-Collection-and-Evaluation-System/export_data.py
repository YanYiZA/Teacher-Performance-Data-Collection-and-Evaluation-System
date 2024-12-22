# -*- coding: utf-8 -*-
import sys
import os
from openpyxl import Workbook
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# 用于处理Excel导出
def export_to_excel(output_file):
    # 创建一个新的Excel工作簿和工作表
    wb = Workbook()
    ws = wb.active

    # 打开data.txt文件并读取内容
    with open('data.txt', 'r', encoding='utf-8') as f:  # 假设data.txt文件在当前目录
        lines = f.readlines()

    # 将txt文件的内容写入Excel
    for row_num, line in enumerate(lines, start=1):
        cells = line.strip().split('\t')  # 按制表符分割每一行
        for col_num, cell in enumerate(cells, start=1):
            ws.cell(row=row_num, column=col_num, value=cell)  # 写入单元格

    # 保存到指定的Excel文件
    wb.save(output_file)
    print(f"Excel file saved: {output_file}")

# 用于处理Word导出
def export_to_word(output_file):
    # 创建一个新的Word文档
    doc = Document()

    # 预定义表头
    columns = ["Name", "Teaching Work", "Research Work", "Scientific Research", "Other Work", "Total Performance"]

    # 打开data.txt文件并读取内容
    with open('data.txt', 'r', encoding='utf-8') as f:  # 假设data.txt文件在当前目录
        lines = f.readlines()

    # 创建一个表格，假设数据没有特别复杂的结构
    table = doc.add_table(rows=1, cols=len(columns))

    # 填充表格的表头
    hdr_cells = table.rows[0].cells
    for i, header in enumerate(columns):
        hdr_cells[i].text = header

    # 填充数据行
    for line in lines:
        row_cells = table.add_row().cells
        cells = line.strip().split('\t')  # 按制表符分割每一行
        for i, cell in enumerate(cells):
            row_cells[i].text = cell

    # 保存到指定的Word文件
    doc.save(output_file)
    print(f"Word file saved: {output_file}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python export_data.py <format> <output_file>")
        sys.exit(1)

    # 获取命令行参数
    file_format = sys.argv[1].lower()  # 'excel' 或 'word'
    output_file = sys.argv[2]

    # 确保txt文件存在
    if not os.path.exists('data.txt'):
        print(f"Error: The file data.txt does not exist.")
        sys.exit(1)

    # 根据格式选择处理方式
    if file_format == 'excel':
        export_to_excel(output_file)
    elif file_format == 'word':
        export_to_word(output_file)
    else:
        print("Error: Unsupported format. Please specify 'excel' or 'word'.")
        sys.exit(1)
