import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re
from collections import defaultdict
import os
import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font

def close_window():
    root.destroy()

def browse_file():
    fileName = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if fileName:
        processFile(fileName)

def extract_broker_name(broker):
    """ 提取券商名称，不包含数字代号 """
    if isinstance(broker, str):
        # 去掉数字、空格和英文字母
        broker = re.sub(r'[A-Za-z0-9\s]+', '', broker).strip()
        return broker
    return ""

firmBP = defaultdict(int)
firmSP = defaultdict(int)
firmBS = defaultdict(int)
firmSS =defaultdict(int)
stock_code = ""
current_date = datetime.datetime.now()
formatted_date = current_date.strftime("%Y%m%d")

def processFile(filepath):
    global firmBP, firmSP, firmBS, firmSS, stock_code
    try:
        with open(filepath, 'r', encoding='big5') as file:
            lines = file.readlines()
            stock_code = lines[1].split(',')[1].strip()  # 获取股票代码

        df = pd.read_csv(filepath, encoding="big5", header=1, skiprows=1)
  
        for i, row in df.iterrows():
            broker = extract_broker_name(row['券商'])
            if broker:
                firmBP[broker] += round(float(row['價格']) * float(row['買進股數']) / 1000, 2)
                firmBS[broker] += round(float(row['買進股數']) / 1000, 3)
                firmSP[broker] += round(float(row['價格']) * float(row['賣出股數']) / 1000, 2)
                firmSS[broker] += round(float(row['賣出股數']) / 1000, 3)
            broker1 = extract_broker_name(row['券商.1'])
            if broker1:
                firmBP[broker1] += round(float(row['價格.1']) * float(row['買進股數.1']) / 1000, 2)
                firmBS[broker1] += round(float(row['買進股數.1']) / 1000, 3)
                firmSP[broker1] += round(float(row['價格.1']) * float(row['賣出股數.1']) / 1000, 2)
                firmSS[broker1] += round(float(row['賣出股數.1']) / 1000, 3)
        for j in firmBP.keys():
            if firmBS[j] != 0:
                firmBP[j] = round(firmBP[j] / firmBS[j], 2)
                firmBS[j] = round(firmBS[j], 3)
            else:
                firmBP[j] = 0

        for k in firmSP.keys():

            if firmSS[k] != 0:
                firmSP[k] = round(firmSP[k] / firmSS[k], 2)
                firmSS[k] = round(firmSS[k], 3)
            else:
                firmSP[k] = 0
        
        write_to_excel()
        close_window()
    except Exception as e:
        messagebox.showerror("錯誤", f"處裡錯誤: {e}")

def write_to_excel():
    # 将字典转换为 DataFrame
    df_bp = pd.DataFrame(list(firmBP.items()), columns=['券商', '買入價格'])
    df_bs = pd.DataFrame(list(firmBS.items()), columns=['券商', '買入張數'])
    df_sp = pd.DataFrame(list(firmSP.items()), columns=['券商', '賣出價格'])
    df_ss = pd.DataFrame(list(firmSS.items()), columns=['券商', '賣出張數'])

    # 合并买入和卖出数据
    df_buy = pd.merge(df_bp, df_bs, on='券商', how='outer')
    df_sell = pd.merge(df_sp, df_ss, on='券商', how='outer')

    # 合并买卖数据
    df_all = pd.merge(df_buy, df_sell, on='券商', how='outer', suffixes=('_買入', '_賣出')).fillna(0)

    # 按买入张数排序
    df_all = df_all.sort_values(by='買入張數', ascending=False)

    # 计算盈虧(萬)
    df_all['盈虧(萬)'] = round(((df_all['賣出價格'] * df_all['賣出張數']) - (df_all['買入價格'] * df_all['買入張數'])) / 10, 1)
    df_all['買賣超(張)'] = df_all['買入張數'] - df_all['賣出張數']

    # 移动券商列到第三列
    broker_column = df_all.pop('券商')
    df_all.insert(2, '券商', broker_column)

    # 创建一个ExcelWriter对象
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    output_file = os.path.join(desktop, f'{stock_code}_{formatted_date}.xlsx')
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_all.to_excel(writer, sheet_name='買賣超', index=False, startrow=1, startcol=1)
    
     # 调整列宽
    workbook = load_workbook(output_file)
    sheet = workbook['買賣超']

    for col in sheet.columns:
        column = col[0].column_letter  # 获取列字母
        sheet.column_dimensions[column].width = 15
    
    # 设置盈虧(萬)列的颜色
    for cell in sheet['G']: 
        if cell.row == 1:  # 跳过标题行
            continue
        try:
            value = float(cell.value)
            if value < 0:
                cell.font = Font(color="FF0000")  # 红色字体
        except (ValueError, TypeError):
            continue
    
    for cell in sheet['H']: 
        if cell.row == 1:  # 跳过标题行
            continue
        try:
            value = float(cell.value)
            if value < 0:
                cell.font = Font(color="FF0000")  # 红色字体
        except (ValueError, TypeError):
            continue

    workbook.save(output_file)
    
    messagebox.showinfo("訊息", f"當日數據寫入Excel文件: {output_file}")

root = tk.Tk()
root.title("成交均價計算")

# 设置窗口大小
window_width = 300
window_height = 200

# 获取屏幕尺寸以计算窗口居中的位置
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# 计算窗口左上角的位置
position_top = int(screen_height / 2 - window_height / 2)
position_right = int(screen_width / 2 - window_width / 2)

# 设置窗口的大小和位置
root.geometry(f"{window_width}x{window_height}+{position_right}+{position_top}")

# 创建一个框架来容纳按钮
frame = tk.Frame(root)
frame.grid(row=2, column=0, pady=20)

# 添加文字标签
label1 = tk.Label(frame, text="加入證交所個股交易量價CSV檔")
label1.grid(row=0, column=0, columnspan=2, pady=(0, 10))

# 添加第二个文字标签
label2 = tk.Label(frame, text="This calculator-V3 powered by Michael")
label2.grid(row=1, column=0, columnspan=2, pady=(0, 10))

# 创建浏览按钮
button = tk.Button(frame, text="瀏覽檔案", command=browse_file)
button.grid(row=2, column=0, padx=10)

# 创建离开按钮
exit = tk.Button(frame, text="離開", command=close_window)
exit.grid(row=2, column=1, padx=10)

# 调整列和行的权重，使按钮框架在窗口中居中
root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=0)
root.grid_rowconfigure(2, weight=1)
root.grid_columnconfigure(0, weight=1)

root.mainloop()

