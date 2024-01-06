# excel_log
PythonプログラムでExcelにログするためのライブラリを作ってみました。ライブラリーとサンプルプログラムをココに置きます。

## 各ファイルの説明
### excel_log.py
　ライブラリー用Pythonプログラムファイル
###excel_test.py
　テスト投稿用Pythonプログラムファイル

## 保存するExcelファイルの設定
ディレクトリ名、ファイル名、シート名、ヘッダー名の設定方法はexcel_log.pyで行います。

### ディレクトリ名とファイル名
#### ディレクトリ名
self, log_dir="C:\\excel_log"
#### ファイル名
log_filename="work_log.xlsx"

```
def __init__(self, log_dir="C:\\excel_log", log_filename="work_log.xlsx", headers=None):
```
### シート名とヘッダー名
```
        self.headers = headers if headers is not None else  {
            "Sheet1": ["Datetime", "Header 1", "Header 2", "Header 3", "Header 4", "Header 5"],
            "Sheet2": ["Datetime", "Header 1", "Header 2", "Header 3", "Header 4", "Header 5"]
            }
```
## インポートライブラリー

```
import os
from openpyxl import Workbook, load_workbook
from datetime import datetime
```

## ログ時間
ログした時間をA列に記録するようにしています。
表示形式は年は西暦で表示し月日時間は24時間制で表示し分と秒あと1/100秒まで表示できるようにしています。
```
current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-4]
```

## サンプルプログラム
テスト投稿用Pythonプログラムexcel_test.pyファイルを使ってExcelにログするコマンドの説明をします。

### ライブラリー
同じフォルダーにある状態での説明です。
ファイル名：excel_log
クラス名：ExcelLogger
```
from excel_log import ExcelLogger
```

### ログするデータ
辞書形式で変数に代入します。

```
data_to_log = {
    "Header 1": "xxxx",
    "Header 2": "oooo",
    "Header 3": "oxoxo",
    "Header 4": "□□□□",
    "Header 5": "xxxxx"
}
```
### コマンド

　データをExcelファイルにログとして記録
 ```
logger.log_to_excel(data_to_log)
```





