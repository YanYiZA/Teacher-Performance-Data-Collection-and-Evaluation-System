import sys
import pandas as pd
from docx import Document

def export_to_excel(data, file_name="output.xlsx"):
    """
    将数据导出到 Excel 文件
    """
    # 创建一个 DataFrame
    df = pd.DataFrame(data[1:], columns=data[0])
    
    # 导出到 Excel
    df.to_excel(file_name, index=False, sheet_name="绩效数据")
    print(f"数据已导出到 Excel 文件：{file_name}")

def export_to_word(data, file_name="output.docx"):
    """
    将数据导出到 Word 文件
    """
    # 创建一个 Word 文档对象
    doc = Document()
    
    # 添加标题
    doc.add_heading('教师绩效数据', 0)
    
    # 创建表格
    table = doc.add_table(rows=1, cols=len(data[0]))
    
    # 添加表头
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(data[0]):
        hdr_cells[i].text = col_name
    
    # 填充数据
    for row in data[1:]:
        row_cells = table.add_row().cells
        for i, cell_value in enumerate(row):
            row_cells[i].text = str(cell_value)
    
    # 保存文档
    doc.save(file_name)
    print(f"数据已导出到 Word 文件：{file_name}")

def parse_data_from_stdin():
    """
    从标准输入获取数据
    """
    data = []
    
    for line in sys.stdin:
        # 处理每一行数据，去掉结尾的换行符并用逗号分隔
        row = line.strip().split(',')
        data.append(row)
    
    return data

def main():
    # 获取命令行参数（格式）
    format_type = sys.argv[1] if len(sys.argv) > 1 else "excel"

    # 从标准输入获取数据
    data = parse_data_from_stdin()
    
    # 根据格式选择导出函数
    if format_type == "excel":
        export_to_excel(data)
    elif format_type == "word":
        export_to_word(data)
    else:
        print("未知的文件格式")

if __name__ == "__main__":
    main()
