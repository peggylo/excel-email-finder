# Excel Email Finder

## 中文說明

### 簡介
Excel Email Finder 是一個用於在多個Excel檔案中搜尋特定email地址的Python工具。它能夠自動檢查資料夾中所有Excel檔案 (.xlsx 和 .xls)，尋找包含指定email的檔案，並報告結果。

### 功能
- 搜尋整個資料夾及其子資料夾中的所有Excel檔案
- 自動識別包含email格式的欄位
- 支援多個工作表(Sheet)的搜尋
- 顯示搜尋結果，包括檔案名稱、工作表名稱和欄位名稱

### 使用方法
1. 修改程式碼中的 `folder_path` 為您要搜尋的資料夾路徑
2. 修改 `target_email` 為您要搜尋的email地址
3. 執行程式
4. 檢視結果

### 需求
- Python 3.6+
- pandas
- openpyxl
- xlrd

## English Description

### Introduction
Excel Email Finder is a Python tool designed to search for specific email addresses across multiple Excel files. It automatically checks all Excel files (.xlsx and .xls) in a folder, identifies which files contain the specified email, and reports the results.

### Features
- Searches all Excel files in a folder and its subfolders
- Automatically identifies columns containing email format data
- Supports multi-sheet Excel files
- Shows search results including file name, sheet name, and column name

### How to Use
1. Modify the `folder_path` in the code to your target folder
2. Change the `target_email` to the email address you want to search for
3. Run the program
4. View the results

### Requirements
- Python 3.6+
- pandas
- openpyxl
- xlrd 
