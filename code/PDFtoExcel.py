import pdfplumber
import pandas as pd

save_path = 'pdftoexcel/PDFtoExcel.xlsx'
res_path = 'pdftoexcel/model01.pdf'

# 三线表格设定尝试
config_dict_three_lines = {"vertical_strategy": "text",
               "horizontal_strategy": "text"}

with pd.ExcelWriter(save_path) as writer:
    with pdfplumber.open(res_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            table = page.extract_table(table_settings=config_dict_three_lines)
            if table:
                df = pd.DataFrame(table[1:], columns=table[0])
                df.to_excel(writer, sheet_name=f'Page {page_number}', index=False)