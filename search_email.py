import os
import re
import pandas as pd
from pathlib import Path

def search_email_in_excel_files(folder_path, target_email):
    # 正則表達式用於辨識email格式
    email_pattern = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')
    
    # 用於存儲結果的列表
    found_files = []
    
    # 遍歷指定資料夾中的所有文件
    print(f"開始搜尋資料夾: {folder_path}")
    excel_count = 0
    
    for file_path in Path(folder_path).glob('**/*'):
        if file_path.suffix.lower() in ['.xlsx', '.xls']:
            excel_count += 1
            try:
                print(f"正在處理文件: {file_path.name}")
                # 讀取Excel文件中的所有sheet
                excel_file = pd.ExcelFile(file_path)
                
                # 遍歷每個sheet
                for sheet_name in excel_file.sheet_names:
                    df = excel_file.parse(sheet_name)
                    
                    # 檢查每一列
                    for column in df.columns:
                        # 將列轉換為字串類型以進行檢查
                        col_data = df[column].astype(str)
                        
                        # 檢查此列是否包含email格式資料
                        if any(col_data.str.match(email_pattern)):
                            # 檢查目標email是否在此列中
                            if any(col_data.str.contains(target_email, case=False, na=False)):
                                found_files.append(str(file_path))
                                print(f"✓ 找到目標email的文件: {file_path.name}, Sheet: {sheet_name}, 欄位: {column}")
                                # 找到後不需要檢查此文件的其他列和sheet
                                break
                    
                    # 如果已經在某個sheet找到，就不需要檢查其他sheet
                    if str(file_path) in found_files:
                        break
                        
            except Exception as e:
                print(f"無法處理文件 {file_path.name}: {e}")
    
    print(f"\n總共處理了 {excel_count} 個Excel檔案")         
    return found_files

if __name__ == "__main__":
    # 直接設定搜尋資料夾和email
    folder_path = r"請寫入你的目標資料夾"
    target_email = "請寫入你要搜尋的email"
    
    print("Excel檔案email搜尋工具")
    print("="*50)
    print(f"搜尋資料夾: {folder_path}")
    print(f"搜尋email: {target_email}")
    print("="*50)
    
    print(f"\n開始在指定資料夾中搜尋 {target_email}...\n")
    results = search_email_in_excel_files(folder_path, target_email)
    
    print(f"\n搜尋完成！找到 {target_email} 的文件數: {len(results)}")
    
    if results:
        print("\n包含目標email的文件:")
        for file in results:
            print(os.path.basename(file))
            print(f"完整路徑: {file}")
    else:
        print("\n沒有找到包含此email的文件")
    
    input("\n按Enter鍵結束程式...") 
