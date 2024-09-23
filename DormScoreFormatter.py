import logging
import os
import sys
import pandas as pd
import argparse
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
import re

def load_and_process_csv_files(folder_path):
    all_data = []
    for file in os.listdir(folder_path):
        if file.startswith('WeekScoreManage_') and file.endswith('.csv'):
            file_path = os.path.join(folder_path, file)
            df = pd.read_csv(file_path, encoding='gbk')
            all_data.append(df)
    
    combined_df = pd.concat(all_data, ignore_index=True)
    processed_df = combined_df[['楼号', '周', '房间', '床位', '总分', '整改意见']].copy()
    processed_df['整改意见'] = processed_df['整改意见'].apply(lambda x: re.sub(r'[^\w\s]', '', str(x)))
    processed_df = processed_df.sort_values(['房间', '床位'])
    
    return processed_df

def create_excel_file(df, output_file):
    wb = Workbook()
    ws = wb.active

    # Set column widths
    for col in ['A', 'B', 'C', 'E', 'F', 'G']:
        ws.column_dimensions[col].width = 6
    ws.column_dimensions['D'].width = 25
    ws.column_dimensions['H'].width = 25
    # Add title and header
    title = f"{df['楼号'].iloc[0]}{df['周'].iloc[0]}"
    ws.merge_cells('A1:H1')
    ws['A1'] = title
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

    ws.merge_cells('A2:H2')
    ws['A2'] = "勤工大队楼层长分队统一意见邮箱：thu.lczh@gmail.com"
    ws['A2'].alignment = Alignment(horizontal='center', vertical='center')

    ws.merge_cells('A3:H3')
    ws['A3'] = "如有疑问请联系学生楼长：xxx@mails.tsinghua.edu.cn 或登陆家园网查询具体成绩"
    ws['A3'].alignment = Alignment(horizontal='center', vertical='center')

    # Add data
    headers = ['房间', '床位', '总分', '整改意见']
    ROW_PER_PAGE = 55
    row = 5
    col = 1
    page = 1
    rows_on_page = 5

    for i, header in enumerate('ABCDEFGH'):
        ws[f'{header}4'] = headers[i % 4]
        ws[f'{header}4'].alignment = Alignment(horizontal='center', vertical='center')
        ws[f'{header}4'].font = Font(name='宋体', size=11)

    for _, data_row in df.iterrows():
        if rows_on_page == ROW_PER_PAGE + 1:
            row = 5 if page == 1 else (page - 1) * ROW_PER_PAGE + 1
            col += 4
            rows_on_page = 5 if page == 1 else 1
            if col > 5:
                page += 1
                row = (page - 1) * ROW_PER_PAGE + 1
                col = 1
                rows_on_page = 1

        for i, header in enumerate(headers):
            if data_row[header] == '' or pd.isna(data_row[header]) or data_row[header] == 'nan':
                ws.cell(row=row, column=col+i, value='')
                logging.warning(f"Empty cell at row {row}, column {col+i}: {header}: {data_row[header]}")
            else:
                ws.cell(row=row, column=col+i, value=data_row[header])
            ws.cell(row=row, column=col+i).alignment = Alignment(horizontal='center', vertical='center')
            font_size = 11 - max(0, (len(str(data_row[header])) - 10) // 2)
            ws.cell(row=row, column=col+i).font = Font(name='宋体', size=font_size)
        row += 1
        rows_on_page += 1

    # Set border for all cells
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin_border
    
    wb.save(output_file)

def main():
    parser = argparse.ArgumentParser(description='Process WeekScoreManage CSV files and create Excel report.')
    parser.add_argument('--folder', default='.', help='Path to the folder containing CSV files')
    args = parser.parse_args()

    df = load_and_process_csv_files(args.folder)
    output_file = f"{df['楼号'].iloc[0]}{df['周'].iloc[0]}.xlsx"
    create_excel_file(df, output_file)
    print(f"Excel file '{output_file}' has been created successfully.")

if __name__ == '__main__':
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    main()
    os.system("pause")
