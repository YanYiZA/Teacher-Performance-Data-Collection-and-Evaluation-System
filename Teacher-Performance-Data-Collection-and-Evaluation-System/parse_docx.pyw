# -*- coding: utf-8 -*-
import docx
import sys

def parse_docx(docx_path):
    doc = docx.Document(docx_path)
    data = []
    for table in doc.tables:
        for row_index, row in enumerate(table.rows):
            # 跳过第一行（标题行）
            if row_index == 0:
                continue
            cells = [cell.text.strip() for cell in row.cells]
            data.append(cells)
    return data

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: python parse_docx.py <path_to_docx_file>")
    else:
        docx_path = sys.argv[1]
        data = parse_docx(docx_path)

        # Write output to stdout in UTF-8
        sys.stdout.buffer.write("\n".join([",".join(row) for row in data]).encode("utf-8"))
