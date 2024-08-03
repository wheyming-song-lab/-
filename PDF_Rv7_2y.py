 # ================================= Packages ===========================================
import os
import re
import shutil
import fitz  
import pandas as pd 
import tkinter as tk
import yolo # 引入yolo,py python檔案，以串接模型預測
import train
from tkinter import filedialog
from pdf2image import convert_from_path 
import multiprocessing
multiprocessing.freeze_support()
# ================================= 文件辨識主程式 =================================================
def PDF_READ(file_path, directory, directory2,directory3, output_folder, final_output_folder):

    

    # 檔案整理 開始
    # ================================= PL及藍圖檔案整理 ===========================================

    # PDF_FILES 開始
    
    
    #Part List 整理開始
    def copy_pdf_files(src_folder, dest_folder):
        """遞迴檢查並複製符合條件的PDF文件。"""
        for root, dirs, files in os.walk(src_folder):
            for filename in files:
                if "~PL" in filename and filename.lower().endswith('.pdf'):
                    target_folder_name = filename.split('~')[1]
                    target_folder_path = os.path.join(dest_folder, target_folder_name)
                    os.makedirs(target_folder_path, exist_ok=True)
                    src_file = os.path.join(root, filename)
                    dest_file = os.path.join(target_folder_path, filename)
                    shutil.copy(src_file, dest_file)
                    print(f"PL檔案 {filename} 已複製到 {target_folder_path}")
    # Part List 整理結束

    # MPPL檔案整理 開始
                elif "~MPPL" in filename and filename.lower().endswith('.pdf'):
                    target_folder_name = filename.split('~')[1]
                    target_folder_path = os.path.join(dest_folder, f"{target_folder_name}_MPPL")
                    os.makedirs(target_folder_path, exist_ok=True)
                    src_file = os.path.join(root, filename)
                    dest_file = os.path.join(target_folder_path, filename)
                    shutil.copy(src_file, dest_file)
                    print(f"MPPL檔案 {filename} 已複製到 {target_folder_path}")

    # MPPL檔案整理 結束

    # PDF_FILES 結束

    # 藍圖檔案整理 開始
    def find_and_organize_files(root_folder, prefix="PSE~", file_extension=".tif"):
        for root, dirs, files in os.walk(root_folder):
            for file in files:
                if file.startswith(prefix) and file.endswith(file_extension):
                    parts = file.split('~')
                    if len(parts) > 1 and parts[1]:
                        new_folder_name = parts[1]
                        new_folder_path = os.path.join(directory3, new_folder_name)
                        os.makedirs(new_folder_path, exist_ok=True)
                        file_path = os.path.join(root, file)
                        shutil.copy(file_path, new_folder_path)
                        print(f"檔案 {file} 已經被複製到 {new_folder_path}")

    find_and_organize_files(directory)

    # 藍圖檔案整理 結束

    # PDF&藍圖資料夾建立 開始

    # 擷取主資料夾中所有子資料夾的名稱
    subfolders = [f.name for f in os.scandir(directory) if f.is_dir()]

    # 遍歷每個子資料夾，將其名稱當作新資料夾路徑
    for subfolder in subfolders:
        full_path = os.path.join(directory, subfolder)
        copy_pdf_files(full_path, directory2)

    # PDF&藍圖資料夾建立 結束

    # 檔案整理 結束 






# =================================================-X~PL文件辨識=========================================================
    print("開始處理...")
    print(f"原始BOM檔案: {file_path}")
    print(f"原始檔案位置:{directory}")
    print(f"Part List 檔案: {directory2}")
    print(f"輸出文件辨别结果的文件夾: {output_folder}")
    print(f"最终BOM表輸出位置: {final_output_folder}")
    df = pd.read_excel(file_path, sheet_name='REF MAKE')
    df.fillna(value=0, inplace=True)
    filtered_df = df[df['製程'] == '白鐵']

     # ================================= 讀取BOM上之件號，從整理後之資料夾中找到對應PL檔案 ===========================================


    #-X_PL判斷 開始

    def find_files(directory2, filename):
        """在指定目錄下查找包含特定部分文件名的文件，返回第一個匹配的完整路徑。"""
        for root, dirs, files in os.walk(directory2):
            for file in files:
                if '~' in file:
                    parts = file.split('~')
                    if len(parts) > 1 and filename in parts[1] and "MPPL" not in file:
                        input_folder = os.path.join(root, file)
                        print(f"找到PL文件：{input_folder}")
                        return input_folder
                    

        return None
    
    #-X_PL判斷 結束


    #-X_MPPL判斷 開始
    def find_files_mppl(directory2, filename):
        """在指定目錄下查找包含特定部分文件名的文件，返回第一個匹配的完整路徑。"""
        for root, dirs, files in os.walk(directory2):
            for file in files:
                if '~' in file:
                    parts = file.split('~')
                    if len(parts) > 1 and filename in parts[1] and "MPPL" in file:
                        input_folder2 = os.path.join(root, file)
                        print(f"找到MPPL文件：{input_folder2}")
                        return input_folder2
                    

        return None   
    
    #-X_MPPL判斷 結束



    for num_for_count, filename in enumerate(filtered_df['件號\nPart Number']):

        input_folder = find_files(directory2, filename) #此程式用於找出我的目標Part List 檔案
        input_folder2=find_files_mppl(directory2, filename)

        # 轉txt 開始
        # ================================= PL檔案從PDF資料格式轉乘可編譯txt ================================================
        if input_folder :
            doc = fitz.open(input_folder) #利用PyMuPDF套件中的fitz函式對PDF文件做文字提取
            text_output_path = os.path.join(output_folder, os.path.basename(input_folder).split('~')[1], 'output_txt') # 指定txt檔輸出位置
            if not os.path.exists(text_output_path): # 如果上列指定之位置不存在的話就建立新的路徑
                os.makedirs(text_output_path)
            
            for page_num, page in enumerate(doc):# 遍歷每一頁 PDF
                text = page.get_text() #獲取每一頁PDF的文字內容
                output_txt_file = os.path.join(text_output_path, f"page_{page_num + 1}.txt") # 輸出並指定每一頁TXT檔輸出位置
                with open(output_txt_file, 'w', encoding='UTF-8') as f: 
                    f.write(text) # 將fitz所提取的文字內容輸入至output_txt_file中
            doc.close()

        # 轉txt 結束

        
            
        # 合併txt 開始

            def merge_text_files(folder_path, output_file_path): 
                with open(output_file_path, 'w', encoding='UTF-8') as output_file:
                    for file_name in os.listdir(folder_path):
                        if file_name.endswith('.txt'): # 找出資料夾內為TXT檔的檔案
                            file_path = os.path.join(folder_path, file_name)
                            with open(file_path, 'r', encoding='UTF-8') as input_file:
                                content = input_file.read()
                                # content = content.replace('=', '-').replace('—', '-') (此行原是用於OCR)
                                output_file.write(content + '\n')
                print(f"The text files have been merged and processed into {output_file_path}")

            
            merge_text_files(text_output_path, os.path.join(text_output_path, 'merged_file.txt'))

        # 合併txt 結束


             # =================================提取~PL上的關鍵資訊  ================================================
            def extract_info_and_update_excel(folder_path, df):         
                    for new_col in ['Material Part Num','Material Condition','Material DEFINITION','Material GENERAL CATEGORY','Material SPECIFICATION','Material USE','原始厚度', '原始寬度', '原始長度', 'SHEET Content', '特殊備註','表處內容', '製程內容','原始厚度_MPPL', '原始寬度_MPPL', '原始長度_MPPL', 'SHEET Content_MPPL', '特殊備註_MPPL']: # 在BOM表上新增欲提取的關鍵資訊欄位
                        if new_col not in df.columns:
                            df[new_col] = None 
                    
                    processed = set()
                    for file_name in os.listdir(folder_path):
                        if file_name.startswith('merged'): # 以合併後的TXT檔來進行關鍵資訊提取
                            file_path = os.path.join(folder_path, file_name)
                            with open(file_path, 'r', encoding='utf-8-sig') as file:
                                file_content = file.read()
                                dataset_filename=filename
                                print("件號:",filename)
                                
                                
                                def extract_and_format_dimension(pattern): # 提取關鍵資訊中零件原始尺寸的程式
                                    match = re.search(pattern, file_content, re.IGNORECASE) # 利用正則表達式來做原始尺寸提取，pattern為欲提取的關鍵字邏輯，
                                    if match:
                                    
                                        value = match.group(1).replace(" ", "") 
                                        if value.startswith('.'): 
                                            value = '0' + value  
                                        try:
                                        
                                            return str(float(value))
                                        except ValueError:
                                        
                                            return "None"
                                    return "None"
                                

                                #搜尋材質件號 開始

                                def extract_material(file_content):
                                    material_pattern = r'QTY REQD\s+PART OR IDENTIFYING\s+DESCRIPTION\s+NOTE TITLE\s+NOTE DESCRIPTION\s+\n\s*NUMBER\s*\n\s*\d+\s+([A-Za-z0-9]+)'
                                    matches = re.findall(material_pattern, file_content, re.DOTALL)
                                    for match in matches:
                                        if not re.search(r'[()]', match):
                                            return match.strip()
                                    return None
                                def extract_material_whilenone(file_content):
                                
                                    pattern_step1 = r'QTY REQD\s+PART OR IDENTIFYING\s+DESCRIPTION\s+NOTE TITLE\s+NOTE DESCRIPTION\s+.*?CONTRACT NUMBER.*?NUMBER\s*\n'
                                    match_step1 = re.search(pattern_step1, file_content, re.DOTALL)
                                    
                                    if match_step1:
                                        relevant_text = file_content[match_step1.end():]
                                        pattern_step2 = r'\n\s*\d+\s+([A-Za-z0-9]+)'
                                        match_step2 = re.search(pattern_step2, relevant_text)
                                        if match_step2:
                                            return match_step2.group(1).strip()
                                    return None
                                def extract_material_whilenone_c2(file_content):
                                    pattern_step1 = r'QTY REQD\s+PART OR IDENTIFYING\s+DESCRIPTION\s+NOTE TITLE\s+NOTE DESCRIPTION\s+.*?CONTRACT NUMBER'
                                    match_step1 = re.search(pattern_step1, file_content, re.DOTALL)
                                    
                                    if match_step1:
                                        relevant_text = file_content[match_step1.end():]
                                        pattern_step2 = r'.*?1\s+([A-Za-z0-9]+)'
                                        match_step2 = re.search(pattern_step2, relevant_text,re.DOTALL)
                                        if match_step2:
                                            return match_step2.group(1).strip()
                                    return None
                              
                                #搜尋材質件號 結束


                                # 在BOM中找到輸入之PDF對應的件號
                                if dataset_filename in df['件號\nPart Number'].values:
                                    row_index = df.index[df['件號\nPart Number'] == dataset_filename].tolist()[0]
                                    material = extract_material(file_content)
                                    print("V1 Material:",material)
                                    if material is None:
                                        material = extract_material_whilenone(file_content)
                                        print("V2 Material:",material)
                                    if material is None:
                                        material = extract_material_whilenone_c2(file_content)
                                        print("V3 Material:",material)   

                                    #零件厚度-X_PL 開始

                                    thickness = extract_and_format_dimension(r'THICKNESS\s+([.\s*\d]+)') # 找到PDF中THICKNESS並提取後面的數值，該值即為原始厚度
                                    
                                    #零件厚度-X_PL 結束  
                                     
                                    #零件寬度-X_PL 開始  
                                                                                             
                                    width = extract_and_format_dimension(r'WIDTH\s+([.\s*\d]+)') # 找到PDF中WIDTH並提取後面的數值，該值即為原始寬度
                                    
                                    #零件長度-X_PL 開始   
                                                                      
                                    length = extract_and_format_dimension(r'LENGTH\s+([.\s*\d]+)') # 找到PDF中LENGTH並提取後面的數值，該值即為原始長度

                                    #零件長度-X_PL 結束   

                                    # 表處與製程內容 開始
                                    pattern_f = r"FINISH - SEQUENCE \d+\s+(F-\d{2}\.\d{2})"
                                    matches_f = re.findall(pattern_f, file_content)
                                    print("Matches for 'F-xx.xx':\n", matches_f)  
                                    
                                    
                                    pattern_process = r"F-\d{2}\.\d{2}\s+(.+?)\s+IN ACCORDANCE"
                                    matches_process = re.findall(pattern_process, file_content, re.DOTALL)
                                    print("Matches for process content:\n", matches_process)  
                                    
                                    
                                    if matches_f:
                                        df['表處內容'] = ', '.join(matches_f)
                                    else:
                                        df['表處內容'] = None
                                    
                                    if matches_process:
                                        df['製程內容'] = ' | '.join(matches_process)
                                    else:
                                        df['製程內容'] = None

                                    # 表處與製程內容 結束

                                    #零件藍圖相關資訊提取 開始

                                    drawing_sheet_pattern = r'DRAWING\s+[A-Za-z0-9-]+\s+SHEET\s+([A-Za-z0-9-]+)' #找到PDF中DRAWING並提取後面的數值，該值即為藍圖檔案編號
                                    sheet_matches = re.findall(drawing_sheet_pattern, file_content, re.IGNORECASE)
                                    sheet_content = '+'.join(sheet_matches) if sheet_matches else 'None' # 若是藍圖檔案編號不只一個的話，在BOM中即會以+字號填入多個編號值
                                    
                                    #零件藍圖相關資訊提取 結束                                    
                                    


                                    #材質PL中的關鍵字提取 開始

                                    input_folder_for_material = find_files(directory2, material)
                                    if input_folder_for_material:
                                            doc = fitz.open(input_folder_for_material) #利用PyMuPDF套件中的fitz函式對PDF文件做文字提取
                                            text_output_path_son = os.path.join(output_folder, os.path.basename(input_folder_for_material).split('~')[1], 'output_txt') # 指定txt檔輸出位置
                                            if not os.path.exists(text_output_path_son): # 如果上列指定之位置不存在的話就建立新的路徑
                                                os.makedirs(text_output_path_son)
                                            
                                            for page_num, page in enumerate(doc):# 遍歷每一頁 PDF
                                                text = page.get_text() #獲取每一頁PDF的文字內容
                                                output_txt_file = os.path.join(text_output_path_son, f"page_{page_num + 1}.txt") # 輸出並指定每一頁TXT檔輸出位置
                                                with open(output_txt_file, 'w', encoding='UTF-8') as f: 
                                                    f.write(text) # 將fitz所提取的文字內容輸入至output_txt_file中
                                            doc.close()
                                        
                                            merge_text_files(text_output_path_son, os.path.join(text_output_path_son, 'merged_file.txt'))

                                            def extract_material_details(file_content):
                                                material_condition_pattern = r'MATERIAL CONDITION\s+(.*)'
                                                material_definition_pattern = r'MATERIAL DEFINITION\s+(.*)'
                                                material_general_category_pattern = r'MATERIAL GENERAL CATEGORY\s+(.*)'
                                                material_specification_pattern = r'MATERIAL SPECIFICATION\s+(.*)'
                                                material_use_pattern = r'MATERIAL USE\s+(.*)'


                                                material_condition = re.search(material_condition_pattern, file_content)
                                                material_definition = re.search(material_definition_pattern, file_content)
                                                material_general_category = re.search(material_general_category_pattern, file_content)
                                                material_specification = re.search(material_specification_pattern, file_content)
                                                material_use = re.search(material_use_pattern, file_content)

                                                material_details = [
                                                    material_condition.group(1).strip() if material_condition else None,
                                                    material_definition.group(1).strip() if material_definition else None,
                                                    material_general_category.group(1).strip() if material_general_category else None,
                                                    material_specification.group(1).strip() if material_specification else None,
                                                    material_use.group(1).strip() if material_use else None,
                                                ]

                                                return material_details if all(material_details) else None
                                            file_path_forson= os.path.join(text_output_path_son, 'merged_file.txt')
                                            with open(file_path_forson, 'r') as file:
                                                file_content = file.read()
                                            material_details = extract_material_details(file_content)



                                            if material_details is not None:                                                                                            
                                                df.at[row_index, 'Material Condition'] = material_details[0]
                                                df.at[row_index, 'Material DEFINITION'] = material_details[1]
                                                df.at[row_index, 'Material GENERAL CATEGORY'] = material_details[2]
                                                df.at[row_index, 'Material SPECIFICATION'] = material_details[3]
                                                df.at[row_index, 'Material USE'] = material_details[4]

                                        #材質PL中的關鍵字提取 結束

                                # 零件特殊製程關鍵字提取 開始
                                    # 檢查文件內容是否有特殊製程，若是存在像是"PEEN"、"PENETRANT"等字樣即輸入對應特徵至BOM表
                                    if "PEEN" in file_content:
                                        outsourcing = "珠擊"
                                    elif "PENETRANT" in file_content or "INSPECT" in file_content:
                                        outsourcing = "滲透檢驗"
                                    else:
                                        outsourcing = "None"  

                                # 零件特殊製程關鍵字提取 結束



                                    # 更新DataFrame，並輸入即更新BOM表
                                    df.at[row_index,'Material Part Num']=material
                                    df.at[row_index, '原始厚度'] = thickness
                                    df.at[row_index, '原始寬度'] = width
                                    df.at[row_index, '原始長度'] = length
                                    df.at[row_index, 'SHEET Content'] = sheet_content
                                    df.at[row_index, '特殊備註'] = outsourcing

                    # 輸出EXCEL BOM表
                    df.to_excel(text_output_path+"/extracted.xlsx", index=False, engine='openpyxl') # 這裡的輸出檔案位置為文件辨識結果，為根據每一件號的暫存辨識資料夾
                    print("Excel文件已根據文本文件更新。")

            extract_info_and_update_excel(text_output_path,df)
            print(f'Data extracted and appended to BOM')

            #檢查所有件號都判斷完成(-X_PL)、 輸出PDF最終BOM 開始

            # 儲存最終的Excel文件，判斷邏輯為BOM表的最後一個ROW資料
            if num_for_count == len(filtered_df['件號\nPart Number']) - 1:
                # final_output_folder = "D:/論文研究/最終bom"
                os.makedirs(final_output_folder, exist_ok=True)  # 確保資料夾存在
                final_output_path = os.path.join(final_output_folder, "最終BOM.xlsx")
                df.to_excel(final_output_path, index=False)
                print(f"最終的BOM已儲存到 {final_output_path}")

            ''' output_csv_path = output_text_base+"/extracted_data.csv"
                extract_info_and_save_to_csv(output_text_base, output_csv_path)
                print(f'Data extracted and saved to {output_csv_path}')'''
            
            #檢查所有件號都判斷完成(-X_PL)、 輸出PDF最終BOM  結束

            # ========================================MPPL===============================================================================
            if input_folder2 != None:
                doc = fitz.open(input_folder2) #利用PyMuPDF套件中的fitz函式對PDF文件做文字提取
                text_output_path = os.path.join(output_folder, os.path.basename(input_folder2).split('~')[1]+"_MPPL", 'output_txt') # 指定txt檔輸出位置
                if not os.path.exists(text_output_path): # 如果上列指定之位置不存在的話就建立新的路徑
                    os.makedirs(text_output_path)
                
                for page_num, page in enumerate(doc):# 遍歷每一頁 PDF
                    text = page.get_text() #獲取每一頁PDF的文字內容
                    output_txt_file = os.path.join(text_output_path, f"page_{page_num + 1}.txt") # 輸出並指定每一頁TXT檔輸出位置
                    with open(output_txt_file, 'w', encoding='UTF-8') as f: 
                        f.write(text) # 將fitz所提取的文字內容輸入至output_txt_file中
                doc.close()
                # =================================合併單一件號PDF的所有頁數的TXT檔  ================================================
                def merge_text_files(folder_path, output_file_path): 
                    with open(output_file_path, 'w', encoding='UTF-8') as output_file:
                        for file_name in os.listdir(folder_path):
                            if file_name.endswith('.txt'): # 找出資料夾內為TXT檔的檔案
                                file_path = os.path.join(folder_path, file_name)
                                with open(file_path, 'r', encoding='UTF-8') as input_file:
                                    content = input_file.read()
                                    # content = content.replace('=', '-').replace('—', '-') (此行原是用於OCR)
                                    output_file.write(content + '\n')
                    print(f"The text files have been merged and processed into {output_file_path}")

                
                merge_text_files(text_output_path, os.path.join(text_output_path, 'merged_file.txt'))

                # =================================提取PART LIST上的關鍵資訊(含"-"件號)  ================================================
                def extract_info_and_update_excel(folder_path, df):  
                        for new_col in ['原始厚度_MPPL', '原始寬度_MPPL', '原始長度_MPPL', 'SHEET Content_MPPL', '特殊備註_MPPL']: # 在BOM表上新增欲提取的關鍵資訊欄位
                            if new_col not in df.columns:
                                df[new_col] = None 
                        
                        processed = set()
                        for file_name in os.listdir(folder_path):
                            if file_name.startswith('merged'): # 以合併後的TXT檔來進行關鍵資訊提取
                                file_path = os.path.join(folder_path, file_name)
                                with open(file_path, 'r', encoding='utf-8-sig') as file:
                                    file_content = file.read()
                                    dataset_filename=filename
                                    print("件號:",filename)
                                    
                                    
                                    def extract_and_format_dimension(pattern): # 提取關鍵資訊中零件原始尺寸的程式
                                        match = re.search(pattern, file_content, re.IGNORECASE) # 利用正則表達式來做原始尺寸提取，pattern為欲提取的關鍵字邏輯，
                                        if match:
                                        
                                            value = match.group(1).replace(" ", "") 
                                            if value.startswith('.'): 
                                                value = '0' + value  
                                            try:
                                            
                                                return str(float(value))
                                            except ValueError:
                                            
                                                return "None"
                                        return "None"
                                
                                    # 在BOM中找到輸入之PDF對應的件號
                                    if dataset_filename in df['件號\nPart Number'].values:
                                        row_index = df.index[df['件號\nPart Number'] == dataset_filename].tolist()[0]
                                
                                        thickness = extract_and_format_dimension(r'THICKNESS\s+([.\s*\d]+)') # 找到PDF中THICKNESS並提取後面的數值，該值即為原始厚度
                                        width = extract_and_format_dimension(r'WIDTH\s+([.\s*\d]+)') # 找到PDF中WIDTH並提取後面的數值，該值即為原始寬度
                                        length = extract_and_format_dimension(r'LENGTH\s+([.\s*\d]+)') # 找到PDF中LENGTH並提取後面的數值，該值即為原始長度

                                        drawing_sheet_pattern = r'DRAWING\s+[A-Za-z0-9-]+\s+SHEET\s+([A-Za-z0-9-]+)' #找到PDF中DRAWING並提取後面的數值，該值即為藍圖檔案編號
                                        sheet_matches = re.findall(drawing_sheet_pattern, file_content, re.IGNORECASE)
                                        sheet_content = '+'.join(sheet_matches) if sheet_matches else 'None' # 若是藍圖檔案編號不只一個的話，在BOM中即會以+字號填入多個編號值
                                        
                                        # 檢查文件內容是否有特殊製程，若是存在像是"PEEN"、"PENETRANT"等字樣即輸入對應特徵至BOM表
                                        if "PEEN" in file_content:
                                            outsourcing = "珠擊"
                                        elif "PENETRANT" in file_content or "INSPECT" in file_content:
                                            outsourcing = "滲透檢驗"
                                        else:
                                            outsourcing = "None"  
                                        # 更新DataFrame，並輸入即更新BOM表
                                        df.at[row_index, '原始厚度_MPPL'] = thickness
                                        df.at[row_index, '原始寬度_MPPL'] = width
                                        df.at[row_index, '原始長度_MPPL'] = length
                                        df.at[row_index, 'SHEET Content_MPPL'] = sheet_content
                                        df.at[row_index, '特殊備註_MPPL'] = outsourcing

                        # 輸出EXCEL BOM表
                        df.to_excel(text_output_path+"/extracted.xlsx", index=False, engine='openpyxl') # 這裡的輸出檔案位置為文件辨識結果，為根據每一件號的暫存辨識資料夾
                        print("Excel文件已根據文本文件更新。")
                extract_info_and_update_excel(text_output_path,df)
                print(f'Data extracted and appended to BOM')

                # 儲存最終的Excel文件，判斷邏輯為BOM表的最後一個ROW資料
                if num_for_count == len(filtered_df['件號\nPart Number']) - 1:
                    # final_output_folder = "D:/論文研究/最終bom"
                    os.makedirs(final_output_folder, exist_ok=True)  # 確保資料夾存在
                    final_output_path = os.path.join(final_output_folder, "最終BOM.xlsx")
                    df.to_excel(final_output_path, index=False)
                    print(f"最終的BOM已儲存到 {final_output_path}")            

        # =================================提取PART LIST上的關鍵資訊(含"-"件號:件號名稱為主件號搭配一子件號)  ================================================    
        else:
            print(f"在 {directory2} 中未找到：{filename}，可能為非獨立檔案")
            filename_for_count=filename.split('-')
            if len(filename_for_count)==2: # 在此以'-'來做件號名稱區隔，計算區隔後的字串數量，若數量為2即為主件號搭配一子件號的格式
                filename1=filename.split('-')[0]
                # output_folder="D:/OCR/train_image/TIME_COUNT_v6/"
                input_folder = find_files(directory2, filename1)
                
                # ================================= PDF檔案轉檔 ================================================ 
                if input_folder:
                    doc = fitz.open(input_folder)
                    text_output_path = os.path.join(output_folder, os.path.basename(input_folder).split('~')[1], 'output_txt')
                    if not os.path.exists(text_output_path):
                        os.makedirs(text_output_path)
                    
                    for page_num, page in enumerate(doc):
                        text = page.get_text()
                        output_txt_file = os.path.join(text_output_path, f"page_{page_num + 1}.txt")
                        with open(output_txt_file, 'w', encoding='UTF-8') as f:
                            f.write(text)
                    doc.close()

                    '''
                    此段程式與獨立型件號中整合TXT檔的函式一樣，因此直接呼叫前面寫過的函式即可
                    def merge_text_files(folder_path, output_file_path):
                        with open(output_file_path, 'w', encoding='UTF-8') as output_file:
                            for file_name in os.listdir(folder_path):
                                if file_name.endswith('.txt'):
                                    file_path = os.path.join(folder_path, file_name)
                                    with open(file_path, 'r', encoding='UTF-8') as input_file:
                                        content = input_file.read()
                                        content = content.replace('=', '-').replace('—', '-')
                                        output_file.write(content + '\n')
                        print(f"The text files have been merged and processed into {output_file_path}")
                    '''
                
                    merge_text_files(text_output_path, os.path.join(text_output_path, 'merged_file.txt'))


             # =================================PART LIST上的關鍵資訊提取的邏輯撰寫  ================================================
                def extract_info_and_save_to_csv(folder_path, df,file_name):
                    # 新增欄位
                    for new_col in ['SHEET Content', 'ZONE', '原始厚度', '原始長度', '原始寬度']:
                        if new_col not in df.columns:
                            df[new_col] = None
                    
                    for file in os.listdir(folder_path):
                        if file.startswith('merged'):
                            file_path = os.path.join(folder_path, file)
                            with open(file_path, 'r', encoding='utf-8-sig') as file:
                                file_content = file.read()
                                dataset_filename = file_name  
                                print(dataset_filename)
                                
                                if dataset_filename in df['件號\nPart Number'].values:
                                    row_index = df.index[df['件號\nPart Number'] == dataset_filename].tolist()[0]
                                    
                                    # ======================以上程式與含-PL件號撰寫方式相同==================================================

                                    # ======================以下程式為此種形式的零件件號關鍵資訊提取邏輯=======================================

                                    escaped_filename1 = re.escape(dataset_filename.split('-')[1]) # 該行程式提取並定義子件號的數值為何
                                    # zone_pattern = rf"^\s*[\d\-]+\s+-{escaped_filename1}.*?MD\s+ZONE\s+([A-Za-z0-9]+)" # 定義正則表達式關鍵資訊"ZONE"的判斷方式
                                    md_pattern = rf"^\s*[\d\-]+\s+-{escaped_filename1}.*?MD\s+"  
                                    md_pt_pattern = rf"^\s*[\d\-]+\s+-{escaped_filename1}.*?MD\s+PT"  

                                    sheet_content_pattern = rf"^\s*[\d\-]+\s+-{escaped_filename1}.*?DP\s+DRAWING\s+PICTURE\s+SHEET\s+(\d+)" # 定義正則表達式關鍵資訊"零件藍圖編號檔案"的判斷方式
                                    stock_pattern = rf"-{escaped_filename1}.*?STOCK\s+([0-9.]+)\s*X\s*([0-9.]+)\s*X\s*([0-9.]+)|STOCK\s+[^X]+X\s*([0-9.]+)" # 定義正則表達式關鍵資訊"零件原始尺寸"的判斷方式，此處的邏輯依照PL上的零件尺寸呈現形式可分為2種

                                    # ZONE
                                    md_match = re.search(md_pattern, file_content, re.MULTILINE | re.DOTALL)
                                    md_pt_match = re.search(md_pt_pattern, file_content, re.MULTILINE | re.DOTALL)

                                    if md_match:
                                        # MD
                                        md_position = md_match.start()
                                        # 判斷MD後面為ZONE or MD
                                        zone_match = re.search(rf"MD\s+ZONE\s+([A-Za-z0-9]+)", file_content[md_position:], re.MULTILINE | re.DOTALL)
                                        pt_match = re.search(rf"MD\s+PT", file_content[md_position:], re.MULTILINE | re.DOTALL)
                                        
                                        if pt_match and (not zone_match or pt_match.start() < zone_match.start()):
                                            df.at[row_index, 'ZONE'] = None
                                            print(f"MD 後面是 PT，ZONE 設置為 None")
                                        elif zone_match:
                                            df.at[row_index, 'ZONE'] = zone_match.group(1)
                                            print(f"Matched ZONE: {zone_match.group(1)}")

                                    # match = re.search(zone_pattern, file_content, re.MULTILINE | re.DOTALL) # 利用正則表達式找出針對"ZONE"匹配的關鍵數值，該處使用re.MULTILINE | re.DOTALL 原因為子件號與ZONE字樣可能不存在於同一行
                                    # if match:
                                    #     df.at[row_index, 'ZONE'] = match.group(1) 
                                    #     print(f"Matched ZONE: {match.group(1)}")

                                
                                    match = re.search(sheet_content_pattern, file_content, re.MULTILINE | re.DOTALL) # 利用正則表達式找出針對"零件藍圖檔案編號"匹配的關鍵數值，該處使用re.MULTILINE | re.DOTALL 原因為子件號與SHEET字樣可能不存在於同一行
                                    if match:
                                        df.at[row_index, 'SHEET Content'] = match.group(1)
                                        print(f"Matched SHEET Content: {match.group(1)}")

                                    # zone_value = df.at[row_index, 'ZONE']
                                    sheet_content_value = df.at[row_index, 'SHEET Content']

                                    # if zone_value is not None:
                                    #     # 找到第一個非數字字符的索引位置
                                    #     first_non_digit_index = next((i for i, char in enumerate(zone_value) if not char.isdigit()), len(zone_value))
                                    #     # 提取首個數字序列
                                    #     first_group_zone = zone_value[:first_non_digit_index]
                                        
                                    #     if first_group_zone:
                                    #         first_group_zone = int(first_group_zone)
                                        
                                    #     if pd.notna(sheet_content_value):
                                    #         if sheet_content_value != first_group_zone:
                                    #             df.at[row_index, 'ZONE'] = str(int(sheet_content_value)) + zone_value[first_non_digit_index:]
                                    #     else:
                                    #         df.at[row_index, 'SHEET Content'] = first_group_zone


                                    match = re.search(stock_pattern, file_content, re.DOTALL) # 利用正則表達式找出針對"零件原始尺寸"匹配的關鍵數值
                                    if match:
                                        # 依照PL上零件原始尺寸的呈現方式可分為2種，分別為只有長度資訊的PL以及後、寬、長都有的PL
                                        df.at[row_index, '原始厚度'] = match.group(1) if match.group(1) else None
                                        df.at[row_index, '原始長度'] = match.group(2) if match.group(2) else match.group(4)
                                        df.at[row_index, '原始寬度'] = match.group(3) if match.group(3) else None
                                    
                                    # 檢查文件內容是否有特殊製程，若是存在像是"PEEN"、"PENETRANT"等字樣即輸入對應特徵至BOM表
                                    if "PEEN" in file_content:
                                        outsourcing = "珠擊"
                                    elif "PENETRANT" in file_content or "INSPECT" in file_content:
                                        outsourcing = "滲透檢驗"
                                    else:
                                        outsourcing = "None"  

                                    df.at[row_index, '特殊備註'] = outsourcing

                    # 輸出至Excel
                    df.to_excel(folder_path + "/extracted.xlsx", index=False, engine='openpyxl')
                    print("Excel文件已根據文本文件的數據更新。")
                        
                # output_csv_path = output_text_base+"/extracted_data.csv"
                extract_info_and_save_to_csv(text_output_path,df,filename)
                print(f'Data extracted and appended to BOM')

                if num_for_count == len(filtered_df['件號\nPart Number']) - 1:
                    # 儲存最終的Excel文件
                    # final_output_folder = "D:/論文研究/最終bom"
                    os.makedirs(final_output_folder, exist_ok=True)  # 確保資料夾存在
                    final_output_path = os.path.join(final_output_folder, "最終BOM.xlsx")
                    df.to_excel(final_output_path, index=False)
                    print(f"最終的BOM已儲存到 {final_output_path}")


             # =================================提取PART LIST上的關鍵資訊  ================================================
            
            #同樣以計算以'-'切割後的字串數量來進行判斷
            elif len(filename_for_count)==3:
                second_dash_index = filename.find("-", filename.find("-") + 1)
                filename1=filename[:second_dash_index]

                # ================================= PDF檔案轉檔(程式與其他形式的件號撰寫方式相同) ================================================ 

                input_folder = find_files(directory2, filename1)
                if input_folder:
                    doc = fitz.open(input_folder)
                    text_output_path = os.path.join(output_folder, os.path.basename(input_folder).split('~')[1], 'output_txt')
                    if not os.path.exists(text_output_path):
                        os.makedirs(text_output_path)
                    
                    for page_num, page in enumerate(doc):
                        text = page.get_text()
                        output_txt_file = os.path.join(text_output_path, f"page_{page_num + 1}.txt")
                        with open(output_txt_file, 'w', encoding='UTF-8') as f:
                            f.write(text)
                    doc.close()
                    '''
                    def merge_text_files(folder_path, output_file_path):
                        """合併指定文件夾内的所有文本文件"""
                        with open(output_file_path, 'w', encoding='UTF-8') as output_file:
                            for file_name in os.listdir(folder_path):
                                if file_name.endswith('.txt'):
                                    file_path = os.path.join(folder_path, file_name)
                                    with open(file_path, 'r', encoding='UTF-8') as input_file:
                                        content = input_file.read()
                                        content = content.replace('=', '-').replace('—', '-')
                                        output_file.write(content + '\n')
                        print(f"The text files have been merged and processed into {output_file_path}")

                    '''
                    merge_text_files(text_output_path, os.path.join(text_output_path, 'merged_file.txt'))


             # =================================PART LIST上的關鍵資訊提取的邏輯撰寫  ================================================

                #此段程式與 整合型件號:件號名稱為主件號搭配兩個子件號 程式撰寫邏輯相同
                def extract_info_and_save_to_csv(folder_path, df,file_name):
                    # 新增欄位
                    for new_col in ['SHEET Content', 'ZONE', '原始厚度', '原始長度', '原始寬度']:
                        if new_col not in df.columns:
                            df[new_col] = None
                    
                    for file in os.listdir(folder_path):
                        if file.startswith('merged'):
                            file_path = os.path.join(folder_path, file)
                            with open(file_path, 'r', encoding='utf-8-sig') as file:
                                file_content = file.read()
                                dataset_filename = file_name  
                                print(dataset_filename)
                                
                                if dataset_filename in df['件號\nPart Number'].values:
                                    row_index = df.index[df['件號\nPart Number'] == dataset_filename].tolist()[0]
                                    
                                    # ======================以上程式與含-PL件號撰寫方式相同==================================================

                                    # ======================以下程式為此種形式的零件件號關鍵資訊提取邏輯=======================================

                                    escaped_filename1 = re.escape(dataset_filename.split('-')[1]) # 該行程式提取並定義子件號的數值為何
                                    # zone_pattern = rf"^\s*[\d\-]+\s+-{escaped_filename1}.*?MD\s+ZONE\s+([A-Za-z0-9]+)" # 定義正則表達式關鍵資訊"ZONE"的判斷方式
                                    md_pattern = rf"^\s*[\d\-]+\s+-{escaped_filename1}.*?MD\s+"  
                                    md_pt_pattern = rf"^\s*[\d\-]+\s+-{escaped_filename1}.*?MD\s+PT"  

                                    sheet_content_pattern = rf"^\s*[\d\-]+\s+-{escaped_filename1}.*?DP\s+DRAWING\s+PICTURE\s+SHEET\s+(\d+)" # 定義正則表達式關鍵資訊"零件藍圖編號檔案"的判斷方式
                                    stock_pattern = rf"-{escaped_filename1}.*?STOCK\s+([0-9.]+)\s*X\s*([0-9.]+)\s*X\s*([0-9.]+)|STOCK\s+[^X]+X\s*([0-9.]+)" # 定義正則表達式關鍵資訊"零件原始尺寸"的判斷方式，此處的邏輯依照PL上的零件尺寸呈現形式可分為2種

                                    # ZONE
                                    md_match = re.search(md_pattern, file_content, re.MULTILINE | re.DOTALL)
                                    md_pt_match = re.search(md_pt_pattern, file_content, re.MULTILINE | re.DOTALL)

                                    if md_match:
                                        # MD
                                        md_position = md_match.start()
                                        # 判斷MD後面為ZONE or MD
                                        zone_match = re.search(rf"MD\s+ZONE\s+([A-Za-z0-9]+)", file_content[md_position:], re.MULTILINE | re.DOTALL)
                                        pt_match = re.search(rf"MD\s+PT", file_content[md_position:], re.MULTILINE | re.DOTALL)
                                        
                                        if pt_match and (not zone_match or pt_match.start() < zone_match.start()):
                                            df.at[row_index, 'ZONE'] = None
                                            print(f"MD 後面是 PT，ZONE 設置為 None")
                                        elif zone_match:
                                            df.at[row_index, 'ZONE'] = zone_match.group(1)
                                            print(f"Matched ZONE: {zone_match.group(1)}")

                                    # match = re.search(zone_pattern, file_content, re.MULTILINE | re.DOTALL) # 利用正則表達式找出針對"ZONE"匹配的關鍵數值，該處使用re.MULTILINE | re.DOTALL 原因為子件號與ZONE字樣可能不存在於同一行
                                    # if match:
                                    #     df.at[row_index, 'ZONE'] = match.group(1) 
                                    #     print(f"Matched ZONE: {match.group(1)}")

                                
                                    match = re.search(sheet_content_pattern, file_content, re.MULTILINE | re.DOTALL) # 利用正則表達式找出針對"零件藍圖檔案編號"匹配的關鍵數值，該處使用re.MULTILINE | re.DOTALL 原因為子件號與SHEET字樣可能不存在於同一行
                                    if match:
                                        df.at[row_index, 'SHEET Content'] = match.group(1)
                                        print(f"Matched SHEET Content: {match.group(1)}")

                                    # zone_value = df.at[row_index, 'ZONE']
                                    sheet_content_value = df.at[row_index, 'SHEET Content']

                                    # if zone_value is not None:
                                    #     # 找到第一個非數字字符的索引位置
                                    #     first_non_digit_index = next((i for i, char in enumerate(zone_value) if not char.isdigit()), len(zone_value))
                                    #     # 提取首個數字序列
                                    #     first_group_zone = zone_value[:first_non_digit_index]
                                        
                                    #     if first_group_zone:
                                    #         first_group_zone = int(first_group_zone)
                                        
                                    #     if pd.notna(sheet_content_value):
                                    #         if sheet_content_value != first_group_zone:
                                    #             df.at[row_index, 'ZONE'] = str(int(sheet_content_value)) + zone_value[first_non_digit_index:]
                                    #     else:
                                    #         df.at[row_index, 'SHEET Content'] = first_group_zone


                                    match = re.search(stock_pattern, file_content, re.DOTALL) # 利用正則表達式找出針對"零件原始尺寸"匹配的關鍵數值
                                    if match:
                                        # 依照PL上零件原始尺寸的呈現方式可分為2種，分別為只有長度資訊的PL以及後、寬、長都有的PL
                                        df.at[row_index, '原始厚度'] = match.group(1) if match.group(1) else None
                                        df.at[row_index, '原始長度'] = match.group(2) if match.group(2) else match.group(4)
                                        df.at[row_index, '原始寬度'] = match.group(3) if match.group(3) else None
                                    
                                    # 檢查文件內容是否有特殊製程，若是存在像是"PEEN"、"PENETRANT"等字樣即輸入對應特徵至BOM表
                                    if "PEEN" in file_content:
                                        outsourcing = "珠擊"
                                    elif "PENETRANT" in file_content or "INSPECT" in file_content:
                                        outsourcing = "滲透檢驗"
                                    else:
                                        outsourcing = "None"  

                                    df.at[row_index, '特殊備註'] = outsourcing

                    # 輸出至Excel
                    df.to_excel(folder_path + "/extracted.xlsx", index=False, engine='openpyxl')
                    print("Excel文件已根據文本文件的數據更新。")
                        
                # output_csv_path = output_text_base+"/extracted_data.csv"
                extract_info_and_save_to_csv(text_output_path,df,filename)
                print(f'Data extracted and appended to BOM')
                if num_for_count == len(filtered_df['件號\nPart Number']) - 1:
                    # 儲存最終的Excel文件
                    # final_output_folder = "D:/論文研究/最終bom"
                    os.makedirs(final_output_folder, exist_ok=True)  # 確保資料夾存在
                    final_output_path = os.path.join(final_output_folder, "最終BOM.xlsx")
                    df.to_excel(final_output_path, index=False)
                    print(f"最終的BOM已儲存到 {final_output_path}")


# =====================================GUI介面設計====================================================

# 操作者介面輸入 開始

# 製作GUI操作者介面上瀏覽資料夾的功能
def browse_folder(entry_widget):
    folder_path = filedialog.askdirectory()
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, folder_path)

# 製作GUI操作者介面上瀏覽檔案的功能
def browse_file(entry_widget):
    file_path = filedialog.askopenfilename()
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, file_path)



# 製作GUI操作者介面上執行主程式的功能
def execute_program():
    file_path = entry_file_path.get()
    directory = entry_directory.get()
    directory2 = entry_directory2.get()
    directory3 = entry_directory3.get()
    output_folder = entry_output_folder.get()
    final_output_folder = entry_final_output_folder.get()
    process_type = process_type_var.get()
    model_mode=model_mode_var.get()
    train_file_path = entry_train_file_path.get() if model_mode == "訓練模式" else None

    #執行PDF文件關鍵字提取程式
    PDF_READ(file_path, directory, directory2, directory3, output_folder, final_output_folder)

    # 執行yolo.py並傳遞製程類別參數
    # subprocess.run(['python', 'yolo.py', process_type,final_output_folder], check=True)

    if model_mode=="預測模式":
        # 執行yolo.py檔中的main函數，以辨識藍圖上的特徵
        yolo.main(process_type,final_output_folder,directory3)
    elif model_mode=="訓練模式":
        #執行train.py進行藍圖辨識模型訓練
        trainer = train.Trainer(train_file_path,process_type)
        trainer.train()



    label_result.config(text="處理完成！")

def update_mode_options(*args):
    model_mode = model_mode_var.get()
    if model_mode == "訓練模式":
        label_train_file_path.grid(row=8, column=0)
        entry_train_file_path.grid(row=8, column=1)
        button_train_file_path.grid(row=8, column=2)
    else:
        label_train_file_path.grid_remove()
        entry_train_file_path.grid_remove()
        button_train_file_path.grid_remove()

def main():
    root = tk.Tk() #建立GUI介面
    root.title("報價工時自動化介面")

    frame = tk.Frame(root)
    frame.pack(padx=10, pady=10)

    # 定義操作者輸入變數
    global entry_file_path, entry_directory, entry_directory2,entry_directory3, entry_output_folder, entry_final_output_folder,entry_train_file_path

    global label_train_file_path, button_train_file_path

    entry_file_path, entry_directory, entry_directory2, entry_directory3, entry_output_folder, entry_final_output_folder, entry_train_file_path = (
        tk.Entry(frame, width=50), tk.Entry(frame, width=50), tk.Entry(frame, width=50), tk.Entry(frame, width=50), tk.Entry(frame, width=50), tk.Entry(frame, width=50), tk.Entry(frame, width=50)
    )

    labels = ["原始BOM表檔案路徑:", "漢翔檔案資料夾:", "Part List 檔案輸出位置:", "2D藍圖檔案輸出位置", "文件辨識結果位置:", "PDF文件辨識BOM輸出位置:"]
    entries = [entry_file_path, entry_directory, entry_directory2, entry_directory3, entry_output_folder, entry_final_output_folder]
    actions = [browse_file, browse_folder, browse_folder, browse_folder, browse_folder, browse_folder]
    for i, (label, entry, action) in enumerate(zip(labels, entries, actions)):
        tk.Label(frame, text=label).grid(row=i, column=0)
        entry.grid(row=i, column=1)
        button = tk.Button(frame, text="瀏覽...", command=lambda e=entry, a=action: a(e))
        button.grid(row=i, column=2)

    # 建立下拉式卷軸，選擇製程類別
    global process_type_var
    process_type_var = tk.StringVar(value="SM")
    tk.Label(frame, text="選擇欲辨識的零件藍圖特徵:").grid(row=6, column=0)
    process_options = ["SM", "SB", "SH", "SS","Error"]
    process_menu = tk.OptionMenu(frame, process_type_var, *process_options)
    process_menu.grid(row=6, column=1)

    # 建立下拉式卷軸，選擇執行模式
    global model_mode_var
    model_mode_var = tk.StringVar(value="預測模式")
    tk.Label(frame, text="選擇執行模式:").grid(row=7, column=0)
    mode_options = ["預測模式", "訓練模式"]
    mode_menu = tk.OptionMenu(frame, model_mode_var, *mode_options)
    mode_menu.grid(row=7, column=1)

    model_mode_var.trace("w", update_mode_options)

    # 訓練檔案路徑選項
    label_train_file_path = tk.Label(frame, text="訓練檔案路徑:")
    entry_train_file_path = tk.Entry(frame, width=50)
    button_train_file_path = tk.Button(frame, text="瀏覽...", command=lambda: browse_file(entry_train_file_path))

    button_execute = tk.Button(frame, text="執行", command=execute_program)
    button_execute.grid(row=9, column=1, pady=10)

    global label_result
    label_result = tk.Label(frame, text="")
    label_result.grid(row=10, column=1, pady=10)

    root.mainloop()


# 操作者介面輸入 結束

if __name__ == "__main__":
    main()