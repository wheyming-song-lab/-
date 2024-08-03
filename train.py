import ultralytics
import sys
import os
import multiprocessing
from ultralytics import YOLO
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"


class Trainer:
    def __init__(self, train_data_path,process_type):
        self.train_data_path = train_data_path
        self.process_type = process_type

    def train(self):
        model = YOLO('yolov8s.pt')
        results = model.train(
            data=f'{self.train_data_path}',   # 使用傳遞的訓練檔案路徑
            imgsz=640,                   # 輸入影像大小
            epochs=50,                  # 訓練次數
            patience=100,                # 等待訓練次數，無改善提前結束訓練
            batch=8,                     # 批次大小
            project=f'yolov8_for_{self.process_type}',  # 專案名稱
            name='exp01'                 # 訓練實驗名稱
        )

def main(train_data_path):
    if not os.path.isfile(train_data_path):
        print(f"提供的訓練檔案路徑無效: {train_data_path}")
        sys.exit(1)

    trainer = Trainer(train_data_path)
    trainer.train()

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("請提供訓練檔存放位置!")
        sys.exit(2)

    multiprocessing.freeze_support()
    train_data_path = sys.argv[1]
    process_type=sys.argv[2]
    main(train_data_path,process_type)
