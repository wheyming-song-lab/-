import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk


# Goal: 輸出31分類的結果(如SB..)     2024/11/16 update
# Input:(1) 3D藍圖辨識之後的資訊 (2) 製程分類邏輯
class App:
    # 介面設計 開始

    def __init__(self, root):
        self.root = root
        self.root.title("零件製程分類")


       
        #GUI操作介面位置調整
        self.left_frame = tk.Frame(root)
        self.left_frame.pack(side="left", fill="both", expand=True)
        self.right_frame = tk.Frame(root)
        self.right_frame.pack(side="right", fill="both", expand=True)

        self.data = None
        self.filtered_data = None
        self.current_index = 0
        
        self.load_button = ttk.Button(self.left_frame, text="瀏覽 Excel 文件", command=self.load_file)
        self.load_button.pack()

        self.load_image_button = ttk.Button(self.root, text="瀏覽圖片文件", command=self.load_image)
        self.load_image_button.pack()

   
    

        # 顯示 Part Number 
        self.part_label = ttk.Label(self.left_frame, text="")
        self.part_label.pack()
        

        # GUI STEP L1 開始

        # 初始化第一個選項
        self.first_label = ttk.Label(self.left_frame, text="選擇Catia 游標上顯示之圖形結果:")
        self.first_label.pack()
        
        self.first_var = tk.StringVar()
        self.first_options = ["平形四邊圖形", "曲面圖形"]
        self.first_menu = ttk.Combobox(self.left_frame, textvariable=self.first_var, values=self.first_options, state="readonly")
        self.first_menu.pack()
        self.first_menu.bind("<<ComboboxSelected>>", self.show_second_options)

        # GUI STEP L1 結束


        # 存放第二個選項的容器
        self.second_label = None
        self.second_menu = None
        self.second_var = tk.StringVar()

        # 存放第三個選項的容器
        self.third_label = None
        self.third_menu = None
        self.third_var = tk.StringVar()

        # 存放第三個選項的容器
        self.forth_label = None
        self.forth_menu = None
        self.forth_var = tk.StringVar()

        # 存放第三個選項的容器
        self.fifth_label = None
        self.fifth_menu = None
        self.fifth_var = tk.StringVar()

        # 存放第三個選項的容器
        self.sixth_label = None
        self.sixth_menu = None
        self.sixth_var = tk.StringVar()

        # 存放第三個選項的容器
        self.seventh_label = None
        self.seventh_menu = None
        self.seventh_var = tk.StringVar()


        self.canvas = tk.Canvas(self.root, width=1300 ,height=600)
        self.canvas.pack()
        self.scroll_x = tk.Scrollbar(self.root, orient="horizontal", command=self.canvas.xview)
        self.scroll_x.pack(side="bottom", fill="x")
        self.scroll_y = tk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scroll_y.pack(side="right", fill="y")
        
        self.canvas.configure(xscrollcommand=self.scroll_x.set, yscrollcommand=self.scroll_y.set)
       
        
        self.result_label = None
      

    
        # 下一行按鈕
        self.next_button = ttk.Button(self.left_frame, text="下一個件號", command=self.next_part)
        self.next_button.pack()
        self.next_button.config(state="disabled")
    



    def load_file(self):
        # 打開文件對話框選擇文件
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            file_extension = file_path.split('.')[-1]
            if file_extension == 'xlsx':
                # 使用 openpyxl 引擎讀取 .xlsx 文件
                self.data = pd.read_excel(file_path, sheet_name='REF MAKE', engine='openpyxl')
            elif file_extension == 'xls':
                # 使用 xlrd 引擎讀取 .xls 文件
                self.data = pd.read_excel(file_path, sheet_name='REF MAKE', engine='xlrd')
            else:
                # 顯示錯誤信息
                messagebox.showerror("錯誤", "不支持的文件格式")
                return
            
            self.file_path = file_path
            # 檢查必要的欄位是否存在
            if '件號\nPart Number' in self.data.columns and '製程' in self.data.columns and '製程分類' in self.data.columns:
                self.filter_data()
                self.current_index = 0
                self.display_part_number()
            else:
                # 顯示錯誤信息
                messagebox.showerror("錯誤", "Excel 文件中缺少必要的欄位: '件號\nPart Number', '製程' 或 '製程分類'")
    

    # GUI STEP 1 開始

    def filter_data(self):
        self.filtered_data = self.data[self.data['製程'] == '白鐵']
        if self.filtered_data.empty:
            # 顯示提示信息
            messagebox.showinfo("提示", "沒有 '白鐵' 的零件")

    # GUI STEP 1 結束



    # GUI STEP 2 開始
    def display_part_number(self):
        # 顯示當前件號
        if self.current_index < len(self.filtered_data):
            part_number = self.filtered_data.iloc[self.current_index]['件號\nPart Number']
            self.part_label.config(text=f"當前件號: {part_number}")
            self.next_button.config(state="disabled")
        else:
            # 顯示完成信息
            messagebox.showinfo("完成", "所有件號已處理完畢")
    
    # GUI STEP 2 結束


    # 介面設計 結束

    # L2 開始
    def show_second_options(self, event):
        # 根據第一個選項顯示第二個選項
        if self.first_var.get() == "平形四邊圖形":
            second_options = ["外圍為尖角(角鋁)", "其他(板料)"]
        elif self.first_var.get() == "曲面圖形":
            second_options = ["外圍為尖角(角鋁)", "其他(板料)"]
        
        # 動態生成第二個選項
        if self.second_label:
            self.second_label.pack_forget()
        self.second_label = ttk.Label(self.left_frame, text="選擇零件對應特徵(第二層):")
        self.second_label.pack()
        
        if self.second_menu:
            self.second_menu.pack_forget()
        self.second_menu = ttk.Combobox(self.left_frame, textvariable=self.second_var, values=second_options, state="readonly")
        self.second_menu.pack()
        self.second_menu.bind("<<ComboboxSelected>>", self.show_third_options)
    # L2 結束    


    # L3 開始
    def show_third_options(self, event):
        # 根據前兩個選項顯示第三個選項
        if self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "外圍為尖角(角鋁)":
            third_options = ["兩面間存在高低落差", "兩面間無存在高低落差"]
        elif self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "其他(板料)":
            third_options = ["水平平面俯視下存在凹凸區塊", "Catia游標顯示圓柱狀圖形","其他條件不滿足"]
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "外圍為尖角(角鋁)":
            third_options = ["物件長度30吋以上", "物件長度30吋以下"]
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)":
            third_options = ["兩面間存在高低落差", "兩面間無存在高低落差","零件面皆為曲面"]
        else:
            third_options = []
        
        # 動態生成第三個選項
        if self.third_label:
            self.third_label.pack_forget()
        self.third_label = ttk.Label(self.left_frame, text="選擇零件對應特徵(第三層):")
        self.third_label.pack()

        if self.third_menu:
            self.third_menu.pack_forget()
        self.third_menu = ttk.Combobox(self.left_frame, textvariable=self.third_var, values=third_options, state="readonly")
        self.third_menu.pack()
        self.third_menu.bind("<<ComboboxSelected>>", self. process_third_options)

    def process_third_options(self, event):
        # if self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "外圍為尖角(角鋁)" and self.third_var.get()=="兩面間存在高低落差":
        #     self.show_third_options(event)
        # elif self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "外圍為尖角(角鋁)" and self.third_var.get()=="兩面間無存在高低落差":
        #       self.show_third_options(event)
        if self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="零件面皆為曲面":
            self.show_result(event, "SR")
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)" and self.third_var.get() == "兩面間存在高低落差":
            self.show_result(event, "SH") 
        else:
            self.show_forth_options(event)
    # L3 結束

    # L4 開始
    def show_forth_options(self, event):
        # 根據前兩三個選項顯示第四個選項
        if self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "外圍為尖角(角鋁)" and self.third_var.get()=="兩面間存在高低落差":
            forth_options = ["水平平面俯視下存在凹凸區塊", "厚度越來越薄(Taper特徵)","其他"]
        elif self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "外圍為尖角(角鋁)" and self.third_var.get()=="兩面間無存在高低落差":
            forth_options = ["厚度越來越薄(Taper特徵)","其他"]
        elif self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "其他(板料)" and self.third_var.get()=="水平平面俯視下存在凹凸區塊":
            forth_options = ["厚度越來越薄(Taper特徵)","其他"]
        elif self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "其他(板料)" and self.third_var.get()=="Catia游標顯示圓柱狀圖形":
            forth_options = ["存在焊接型圖示","厚度越來越薄(Taper特徵)","水平平面俯視下存在凹凸區塊","其他"]
        elif self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "其他(板料)" and self.third_var.get()=="其他條件不滿足":
            forth_options = ["厚度越來越薄(Taper特徵)","其他"]
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "外圍為尖角(角鋁)"and self.third_var.get()=="物件長度30吋以上":
            forth_options = ["厚度越來越薄(Taper特徵)","其他"]     
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "外圍為尖角(角鋁)"and self.third_var.get()=="物件長度30吋以下":
            forth_options = ["厚度越來越薄(Taper特徵)","Material Noet有寫鈦合金","零件兩面內夾角小於60度","水平平面俯視下存在凹凸區塊","其他"] 
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差":
            forth_options = ["沒有任何面封閉(無frenge)","有任一面有封閉(有frenge)"]   
                          
        else:
            forth_options = []
        
        # 動態生成第四個選項
        if self.forth_label:
            self.forth_label.pack_forget()
        self.forth_label = ttk.Label(self.left_frame, text="選擇零件對應特徵(第四層):")
        self.forth_label.pack()
        
        if self.forth_menu:
            self.forth_menu.pack_forget()
        self.forth_menu = ttk.Combobox(self.left_frame, textvariable=self.forth_var, values=forth_options, state="readonly")
        self.forth_menu.pack()
        self.forth_menu.bind("<<ComboboxSelected>>", self.process_forth_options)


    def process_forth_options(self, event):
        if self.first_var.get() == "曲面圖形" and self.second_var.get() == "外圍為尖角(角鋁)"and self.third_var.get()=="物件長度30吋以上"and self.forth_var.get()=="厚度越來越薄(Taper特徵)":
            self.show_result(event,"SSEM")  
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "外圍為尖角(角鋁)"and self.third_var.get()=="物件長度30吋以上"and self.forth_var.get()=="其他":
            self.show_result(event,"SSE" ) 
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "外圍為尖角(角鋁)"and self.third_var.get()=="物件長度30吋以下"and self.forth_var.get()=="厚度越來越薄(Taper特徵)":
            self.show_result(event,"SPM" )
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "外圍為尖角(角鋁)"and self.third_var.get()=="物件長度30吋以下"and self.forth_var.get()=="Material Noet有寫鈦合金":
            self.show_result(event, "SPH") 
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "外圍為尖角(角鋁)"and self.third_var.get()=="物件長度30吋以下"and self.forth_var.get()=="零件兩面內夾角小於60度":
            self.show_result(event, "SPB") 
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "外圍為尖角(角鋁)"and self.third_var.get()=="物件長度30吋以下"and self.forth_var.get()=="水平平面俯視下存在凹凸區塊":
            self.show_result(event, "SPC")
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "外圍為尖角(角鋁)"and self.third_var.get()=="物件長度30吋以下"and self.forth_var.get()=="其他":
            self.show_result(event,"SP")
        elif self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "其他(板料)" and self.third_var.get()=="其他條件不滿足"and self.forth_var.get()=="厚度越來越薄(Taper特徵)":
            self.show_result(event,"SMM")
        elif self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "其他(板料)" and self.third_var.get()=="其他條件不滿足"and self.forth_var.get()=="其他":
            self.show_result(event,"SM")
        elif self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "其他(板料)" and self.third_var.get()=="Catia游標顯示圓柱狀圖形"and self.forth_var.get()=="存在焊接型圖示":
            self.show_result(event,"SBW")
        elif self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "其他(板料)" and self.third_var.get()=="Catia游標顯示圓柱狀圖形"and self.forth_var.get()=="厚度越來越薄(Taper特徵)":
            self.show_result(event,"SBM")
        elif self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "其他(板料)" and self.third_var.get()=="Catia游標顯示圓柱狀圖形"and self.forth_var.get()=="水平平面俯視下存在凹凸區塊":
            self.show_result(event,"SBC")
        elif self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "其他(板料)" and self.third_var.get()=="Catia游標顯示圓柱狀圖形"and self.forth_var.get()=="其他":
            self.show_result(event,"SB")
        elif self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "其他(板料)" and self.third_var.get()=="水平平面俯視下存在凹凸區塊"and self.forth_var.get()=="厚度越來越薄(Taper特徵)":
            self.show_result(event,"SCM")
        elif self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "其他(板料)" and self.third_var.get()=="水平平面俯視下存在凹凸區塊"and self.forth_var.get()=="其他":
            self.show_result(event,"SC")
        elif self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "外圍為尖角(角鋁)" and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="其他":
            self.show_result(event,"SM")
        elif self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "外圍為尖角(角鋁)" and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="厚度越來越薄(Taper特徵)":
            self.show_result(event,"SMM")
        elif self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "外圍為尖角(角鋁)" and self.third_var.get()=="兩面間存在高低落差"and self.forth_var.get()=="水平平面俯視下存在凹凸區塊":
            self.show_result(event,"SJC")
        elif self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "外圍為尖角(角鋁)" and self.third_var.get()=="兩面間存在高低落差"and self.forth_var.get()=="厚度越來越薄(Taper特徵)":
            self.show_result(event,"SJM")
        elif self.first_var.get() == "平形四邊圖形" and self.second_var.get() == "外圍為尖角(角鋁)" and self.third_var.get()=="兩面間存在高低落差"and self.forth_var.get()=="其他":
            self.show_result(event,"SJ")
        else:
            self.show_fifth_options(event)

    # L4 結束


    # L5 開始
    def show_fifth_options(self,event):
        if self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="沒有任何面封閉(無frenge)":
            fifth_options = ["厚度越來越薄(Taper特徵)","水平平面俯視下存在凹凸區塊","其他"] 
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="有任一面有封閉(有frenge)":
            fifth_options = ["曲面高度4吋以下","曲面高度4吋以上"] 
        # 動態生成第三個選項
        if self.fifth_label:
            self.fifth_label.pack_forget()
        self.fifth_label = ttk.Label(self.left_frame, text="選擇零件對應特徵(第五層):")
        self.fifth_label.pack()

        if self.fifth_menu:
            self.fifth_menu.pack_forget()
        self.fifth_menu = ttk.Combobox(self.left_frame, textvariable=self.fifth_var, values=fifth_options, state="readonly")
        self.fifth_menu.pack()
        self.fifth_menu.bind("<<ComboboxSelected>>", self.process_fifth_options)

    def process_fifth_options(self, event):
        if self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="沒有任何面封閉(無frenge)"and self.fifth_var.get()=="厚度越來越薄(Taper特徵)":
            self.show_result(event,"SSM")  
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="沒有任何面封閉(無frenge)"and self.fifth_var.get()=="其他":
            self.show_result(event,"SS" ) 
        else:
            self.show_sixth_options(event)  

    # L5 結束


    # L6 開始
    def show_sixth_options(self,event):
       if self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="沒有任何面封閉(無frenge)"and self.fifth_var.get()=="水平平面俯視下存在凹凸區塊":
            sixth_options = ["厚度越來越薄(Taper特徵)","其他"] 
       elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="有任一面有封閉(有frenge)"and self.fifth_var.get()=="曲面高度4吋以下":
            sixth_options = ["零件兩面內夾角小於60度","存在焊接型圖示","厚度越來越薄(Taper特徵)","四面封閉","水平平面俯視下存在凹凸區塊","其他"]
       elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="有任一面有封閉(有frenge)"and self.fifth_var.get()=="曲面高度4吋以上":
            sixth_options = ["零件深度高於500絲","水平平面俯視下存在凹凸區塊","其他"] 
        # 動態生成第三個選項
       if self.sixth_label:
            self.sixth_label.pack_forget()
       self.sixth_label = ttk.Label(self.left_frame, text="選擇零件對應特徵(第六層):")
       self.sixth_label.pack()

       if self.sixth_menu:
            self.sixth_menu.pack_forget()
       self.sixth_menu = ttk.Combobox(self.left_frame, textvariable=self.sixth_var, values=sixth_options, state="readonly")
       self.sixth_menu.pack()
       self.sixth_menu.bind("<<ComboboxSelected>>", self.process_sixth_options)  

    def process_sixth_options(self,event):
        if self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="沒有任何面封閉(無frenge)"and self.fifth_var.get()=="水平平面俯視下存在凹凸區塊"and self.sixth_var.get()=="厚度越來越薄(Taper特徵)":
            self.show_result(event,"SSCM") 
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="沒有任何面封閉(無frenge)"and self.fifth_var.get()=="水平平面俯視下存在凹凸區塊"and self.sixth_var.get()=="其他":
            self.show_result(event,"SSC")
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="有任一面有封閉(有frenge)"and self.fifth_var.get()=="曲面高度4吋以下"and self.sixth_var.get()=="零件兩面內夾角小於60度":
            self.show_result(event,"SHB")
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="有任一面有封閉(有frenge)"and self.fifth_var.get()=="曲面高度4吋以下"and self.sixth_var.get()=="存在焊接型圖示":
            self.show_result(event,"SHW")
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="有任一面有封閉(有frenge)"and self.fifth_var.get()=="曲面高度4吋以下"and self.sixth_var.get()=="厚度越來越薄(Taper特徵)":
            self.show_result(event,"SHM")  
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="有任一面有封閉(有frenge)"and self.fifth_var.get()=="曲面高度4吋以下"and self.sixth_var.get()=="水平平面俯視下存在凹凸區塊":
            self.show_result(event,"SHC") 
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="有任一面有封閉(有frenge)"and self.fifth_var.get()=="曲面高度4吋以下"and self.sixth_var.get()=="其他":
            self.show_result(event,"SH")
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="有任一面有封閉(有frenge)"and self.fifth_var.get()=="曲面高度4吋以上"and self.sixth_var.get()=="零件深度高於500絲":
            self.show_result(event,"SDC")
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="有任一面有封閉(有frenge)"and self.fifth_var.get()=="曲面高度4吋以上"and self.sixth_var.get()=="水平平面俯視下存在凹凸區塊":
            self.show_result(event,"SHD")
        elif self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="有任一面有封閉(有frenge)"and self.fifth_var.get()=="曲面高度4吋以上"and self.sixth_var.get()=="其他":
            self.show_result(event,"SD") 
        else:
            self.show_seventh_options(event)     

    # L6 結束


    # L7 開始
    def show_seventh_options(self,event):
       if self.first_var.get() == "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="有任一面有封閉(有frenge)"and self.fifth_var.get()=="曲面高度4吋以下"and self.seventh_var.get()=="四面封閉":
            seventh_options = ["零件深度低於500絲","零件深度高於500絲"] 
        # 動態生成第三個選項
       if self.seventh_label:
            self.seventh_label.pack_forget()
       self.seventh_label = ttk.Label(self.left_frame, text="選擇零件對應特徵(第七層):")
       self.seventh_label.pack()

       if self.seventh_menu:
            self.seventh_menu.pack_forget()
       self.seventh_menu = ttk.Combobox(self.left_frame, textvariable=self.seventh_var, values=seventh_options, state="readonly")
       self.seventh_menu.pack()
       self.seventh_menu.bind("<<ComboboxSelected>>", self.process_seventh_options)  

    def process_seventh_options(self,event):
        if self.first_var.get()== "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="有任一面有封閉(有frenge)"and self.fifth_var.get()=="曲面高度4吋以下"and self.seventh_var.get()=="四面封閉"and self.seventh_var.get()=="零件深度低於500絲":
            self.show_result(event,"SP") 
        elif self.first_var.get()== "曲面圖形" and self.second_var.get() == "其他(板料)"and self.third_var.get()=="兩面間無存在高低落差"and self.forth_var.get()=="有任一面有封閉(有frenge)"and self.fifth_var.get()=="曲面高度4吋以下"and self.seventh_var.get()=="四面封閉"and self.seventh_var.get()=="零件深度高於500絲":
            self.show_result(event,"SHD") 
        else:
            self.show_result(event,"未知分類，須建立至程式資料庫內!!")    

    # L7 結束

    # 輸出類別結果 開始
    def show_result(self, event,result):       
        # 顯示結果
        if self.result_label:
            self.result_label.pack_forget()
        self.result_label = ttk.Label(self.left_frame, text=f"分類項目: {result}")
        self.result_label.pack()
        
        # 更新分類結果並啟用下一個件號按鈕
        self.filtered_data.at[self.filtered_data.index[self.current_index], '製程分類'] = result
        self.data.at[self.filtered_data.index[self.current_index], '製程分類'] = result
        self.save_file()
        self.next_button.config(state="normal")
    
    # 輸出類別結果 開始
    


    #顯示下一件號 開始

    def next_part(self):
        # 顯示下一個件號
        self.current_index += 1
        self.first_var.set('')
        self.second_var.set('')
        self.third_var.set('')
        self.forth_var.set('')
        self.fifth_var.set('')
        self.sixth_var.set('')
        self.seventh_var.set('')
        if self.second_label:
            self.second_label.pack_forget()
        if self.second_menu:
            self.second_menu.pack_forget()
        if self.third_label:
            self.third_label.pack_forget()
        if self.third_menu:
            self.third_menu.pack_forget()
        if self.forth_label:
            self.forth_label.pack_forget()
        if self.forth_menu:
            self.forth_menu.pack_forget()
        if self.fifth_label:
            self.fifth_label.pack_forget()
        if self.fifth_menu:
            self.fifth_menu.pack_forget()
        if self.sixth_label:
            self.sixth_label.pack_forget()
        if self.sixth_menu:
            self.sixth_menu.pack_forget()
        if self.seventh_label:
            self.seventh_label.pack_forget()
        if self.seventh_menu:
            self.seventh_menu.pack_forget()    
        if self.result_label:
            self.result_label.pack_forget()
        self.display_part_number()

    #顯示下一件號 結束


    #輸出類別結果 開始
    def save_file(self):
        if self.file_path.split('.')[0]+'_new.xls' is None:
            os.makedirs(self.file_path.split('.')[0]+'_new.xls')
        self.data.to_excel(self.file_path.split('.')[0]+'_new.xls', index=False, engine='openpyxl')

   #輸出類別結果 結束


    # GUI Step 3 開始
    def load_image(self):
        
        file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg;*.jpeg;*.png;*.bmp;*.tif")])
        if file_path:
            
            self.image = Image.open(file_path)
            self.photo = ImageTk.PhotoImage(self.image)
            
            
            self.canvas.delete("all")
            
            
            self.canvas.create_image(0, 0, anchor="nw", image=self.photo)
            self.canvas.config(scrollregion=self.canvas.bbox("all"))
    # GUI Step 3 結束
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
