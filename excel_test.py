from excel_log import ExcelLogger

# ExcelLogger オブジェクトの作成
# 必要に応じて log_dir, log_filename, headers をカスタマイズしてください
#logger = ExcelLogger(log_dir="C:\\your_log_directory", log_filename="your_log_file.xlsx")
logger = ExcelLogger(log_dir="C:\\excel_log", log_filename="your_log_file.xlsx")

# ログとして記録するデータ
data_to_log = {
    "Header 1": "xxxx",
    "Header 2": "oooo",
    "Header 3": "oxoxo",
    "Header 4": "□□□□",
    "Header 5": "xxxxx"
}

# データをExcelファイルにログとして記録
logger.log_to_excel(data_to_log)
