import os
import sys
import pandas as pd
from ultralytics import YOLO

# 搜尋藍圖 開始
def find_folder(directory, folder_name):
    prefix = folder_name.split('-')[0] 
    for root, dirs, _ in os.walk(directory):
        matched_dirs = [d for d in dirs if d.startswith(prefix)]
        if matched_dirs:
            return os.path.join(root, matched_dirs[0])
        if folder_name in dirs:
            return os.path.join(root, folder_name)
    return None

def find_target_files(directory, part_number, sheet_content):
    target_files = []
    prefix = part_number.split('-')[0]
    
    content_parts = str(sheet_content).split('+')
    for part in content_parts:
        if part.isdigit():  
            
            formatted_number = f"{int(part):04}"  
            pattern = f"~{prefix}~DWG~{formatted_number}~"
            for root, _, files in os.walk(directory):
                for file in files:
                    if pattern in file:
                        target_files.append(os.path.join(root, file))
    return target_files

# 搜尋藍圖 結束


def main(process_type, file_path, directory):
    # 根據製程類別選擇模型
    base_path = os.path.dirname(os.path.abspath(__file__))
    model_paths = {
    "SM": os.path.join(base_path, "models", "SM_best.pt"),
    "SB": os.path.join(base_path, "models", "SB_best.pt"),
    "SH": os.path.join(base_path, "models", "SH_best.pt"),
    "SS": os.path.join(base_path, "models", "SS_best.pt"),
    "Error":os.path.join(base_path,"models","error.pt")
    }

    if process_type not in model_paths:
        print(f"無效的製程類別: {process_type}")
        return
    

    # 讀取件號 開始
    file_path= os.path.join(file_path,"最終BOM.xlsx")
    model_path = model_paths[process_type]
    model = YOLO(model_path) 
    df = pd.read_excel(file_path, sheet_name='Sheet1')
    df.fillna(value=0, inplace=True)
    if process_type=="Error":
        filtered_df=df[df['製程']=='白鐵']
    else:
        filtered_df = df[df['製程分類'] == process_type]

    for _, row in filtered_df.iterrows():
        part_number = row['件號\nPart Number']
        sheet_content = row['SHEET Content']
        
        if sheet_content in ['DL', 'Unknown', ''] or pd.isnull(sheet_content):
            continue

        folder_path = find_folder(directory, part_number)
        if folder_path:
            target_files = find_target_files(folder_path, part_number, sheet_content)
            for file in target_files:
                result = model.predict(
                    source=file,
                    mode="predict",
                    save=True,
                )


  # 讀取件號 結束

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("請提供製程類別、已完成辨識之BOM路徑、藍圖檔案存放位置")
        sys.exit(1)

    process_type = sys.argv[1]
    file_path = sys.argv[2]
    directory = sys.argv[3]
    main(process_type, file_path, directory)
