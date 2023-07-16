# 題庫管理系統

這是一個使用Python和Tkinter GUI套件開發的題庫管理系統。這個系統允許使用者新增、修改和生成測驗題目。


![image](https://github.com/benjamin920101/Test-Bank-Management-System/assets/47590892/dfb85b7b-dfa7-48fe-9058-83d083c9bc5e)
![image](https://github.com/benjamin920101/Test-Bank-Management-System/assets/47590892/12368bc3-fa4c-40cc-b8b8-fcfa341dfdcb)

## 功能

- 新增問題和答案到題庫
- 修改題庫中的問題和答案
- 生成指定數量的測驗題目
- 將題目和答案導出為Word檔案

## 使用方式

1. 安裝必要的套件

   ```
   pip install openpyxl
   pip install python-docx
   ```

2. 執行程式

   ```
   python quiz_manager.py
   ```

3. 程式界面

   - 問題輸入框：輸入問題的文字
   - 答案輸入框：輸入答案的文字
   - 新增按鈕：點擊以新增問題和答案到題庫
   - 修改按鈕：點擊以修改題庫中的問題和答案
   - 題目數量輸入框：輸入欲生成的測驗題目數量
   - 生成測驗按鈕：點擊以生成測驗題目並導出為Word檔案

## 注意事項

- 系統會自動創建一個名為`questions.xlsx`的Excel檔案來存儲題庫資料。
- Excel檔案預設包含一個名為`Sheet1`的工作表，用於存儲問題和答案。
- 每次新增或修改問題和答案後，請點擊"保存修改"按鈕以確保資料已成功保存到Excel檔案。
- 系統會自動生成兩個Word檔案，分別為`quiz.docx`和`answers.docx`，分別包含測驗題目和答案。
- 若無法導出測驗和答案，可能是因為缺少必要的套件或發生了其他錯誤。

請確保在執行程式之前已安裝必要的套件，並在使用系統時按照操作指示進行操作。如有其他問題，請聯繫開發者進行解決。
