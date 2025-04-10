#data log plot and EEF simple calculate for VM7000/PW3335 Data Collection
#-------------------------------------------------------------------------------
#Rev.1.4 新增日期時間增減功能
#-------------------------------------------------------------------------------
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter
import configparser
from matplotlib import rcParams

# 設定 matplotlib 使用的字體
rcParams['font.sans-serif'] = ['Microsoft JhengHei']  # 使用微軟正黑體
rcParams['axes.unicode_minus'] = False  # 解決負號無法顯示的問題

# 初始化主視窗
root = tk.Tk()
root.title("CSV Plotter with Statistics 1.4")

# 全域變數
csv_path = tk.StringVar()
chart_title = tk.StringVar()
start_datetime = tk.StringVar()
end_datetime = tk.StringVar()
start_date = tk.StringVar()
start_time = tk.StringVar()
end_date = tk.StringVar()
end_time = tk.StringVar()

# 選擇檔案的函數
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        csv_path.set(file_path)
        try:
            # 讀取 CSV 檔案
            df = pd.read_csv(file_path)
            
            # 檢查是否有 Date 和 Time 欄位
            if 'Date' in df.columns and 'Time' in df.columns:
                # 合併 Date 和 Time 欄位為 datetime
                df['datetime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'])
                
                # 設定開始和結束日期時間
                start_date.set(df['datetime'].iloc[0].strftime('%Y-%m-%d'))
                start_time.set(df['datetime'].iloc[0].strftime('%H:%M'))
                end_date.set(df['datetime'].iloc[-1].strftime('%Y-%m-%d'))
                end_time.set(df['datetime'].iloc[-1].strftime('%H:%M'))
            else:
                messagebox.showerror("錯誤", "CSV 檔案缺少 Date 或 Time 欄位！")

            # 檢查是否有 U(V), I(A), P(W), WP(Wh) 欄位，若無則補上並填入 0
            required_columns = ['U(V)', 'I(A)', 'P(W)', 'WP(Wh)']
            for col in required_columns:
                if col not in df.columns:
                    df[col] = 0

        except Exception as e:
            messagebox.showerror("錯誤", f"處理 CSV 檔案時發生錯誤：{e}")

# 儲存設定的函數
def save_config():
    config = configparser.ConfigParser()
    config['Settings'] = {
        "csv_path": csv_path.get(),
        "chart_title": chart_title.get()
    }
    with open("config.ini", "w") as configfile:
        config.write(configfile)
    messagebox.showinfo("設定儲存", "設定已成功儲存！")

# 讀取設定的函數
def load_config():
    try:
        file_path = filedialog.askopenfilename(filetypes=[("INI files", "*.ini")])
        if not file_path:
            return
        config = configparser.ConfigParser()
        config.read(file_path)
        csv_path.set(config['Settings'].get("csv_path", ""))
        chart_title.set(config['Settings'].get("chart_title", ""))
        messagebox.showinfo("設定載入", "設定已成功載入！")
    except FileNotFoundError:
        messagebox.showerror("錯誤", "找不到設定檔！")
    except Exception as e:
        messagebox.showerror("錯誤", f"載入設定時發生錯誤：{e}")

# 繪製圖表的函數
def plot_chart():
    ax1_color = "lightcyan"
    ax2_color = "lightyellow"
    try:
        # 讀取 CSV 檔案
        df = pd.read_csv(csv_path.get())
        
        # 合併 Date 和 Time 欄位為 datetime
        df['datetime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'])
        
        # 檢查是否有 U(V), I(A), P(W), WP(Wh) 欄位，若無則補上並填入 0
        required_columns = ['U(V)', 'I(A)', 'P(W)', 'WP(Wh)']
        for col in required_columns:
            if col not in df.columns:
                df[col] = 0

        # 確認電力欄位是否存在 (與自動補0 重複,暫時保留)
        power_columns = ['U(V)', 'I(A)', 'P(W)']
        if all(col in df.columns for col in power_columns):
            df_power = df[['datetime'] + power_columns]
        else:
            messagebox.showerror("錯誤", "缺少電力欄位！")
            return
        
        # 其餘欄位視為溫度資料
        temp_columns = [col for col in df.columns if col not in ['Date', 'Time', 'datetime', 'WP(Wh)'] + power_columns]
        df_temp = df[['datetime'] + temp_columns]
        
        # 繪製圖表
        fig, (ax1, ax2) = plt.subplots(2, 1, sharex=True, figsize=(10, 6), facecolor='lightgray')
        fig.suptitle(chart_title.get())
        
        # 繪製溫度資料
        for col in temp_columns:
            ax1.plot(df_temp['datetime'], df_temp[col], label=col)
        ax1.set_ylabel("溫度")
        ax1.legend()
        ax1.set_facecolor(ax1_color)
        ax1.grid(True, linestyle='dashdot')
        
        # 繪製電力資料
        ax2a = ax2  # 左側 Y 軸 (功率 P(W))
        ax2b = ax2.twinx()  # 右側第一個 Y 軸 (電壓 U(V))
        ax2c = ax2.twinx()  # 右側第二個 Y 軸 (電流 I(A))
        
        # 調整右側第二個 Y 軸的位置
        ax2c.spines['right'].set_position(('outward', 60))
        
        # 繪製 'P(W)' 資料在左側 Y 軸
        ax2a.plot(df_power['datetime'], df_power['P(W)'], label='P(W)', color='green')
        ax2a.set_ylabel("功率 (P)", color='green')
        ax2a.tick_params(axis='y', labelcolor='green')
        ax2a.grid(True, linestyle='dashdot')
        
        # 繪製 'U(V)' 資料在右側第一個 Y 軸
        ax2b.plot(df_power['datetime'], df_power['U(V)'], label='U(V)', color='blue')
        ax2b.set_ylabel("電壓 (U)", color='blue')
        ax2b.tick_params(axis='y', labelcolor='blue')
        ax2b.set_ylim(0, 120)  # 設定 Y 軸範圍為 0 到 120
        
        # 繪製 'I(A)' 資料在右側第二個 Y 軸
        ax2c.plot(df_power['datetime'], df_power['I(A)'], label='I(A)', color='red')
        ax2c.set_ylabel("電流 (I)", color='red')
        ax2c.tick_params(axis='y', labelcolor='red')
        
        # 設定背景顏色
        ax2.set_facecolor(ax2_color)
        
        # 設定 X 軸格式
        ax2.xaxis.set_major_formatter(DateFormatter("%m-%d %H:%M"))
        ax2.xaxis.set_major_locator(plt.MaxNLocator(10))  # 限制 X 軸最多顯示 15 個標籤
        plt.setp(ax2.xaxis.get_majorticklabels(), rotation=45)
        plt.tight_layout()
        plt.show()
    except Exception as e:
        messagebox.showerror("錯誤", f"繪製圖表時發生錯誤：{e}")

# 計算平均值的函數
def calculate_statistics():
    try:
        # 讀取 CSV 檔案
        df = pd.read_csv(csv_path.get())
        
        # 合併 Date 和 Time 欄位為 datetime
        df['datetime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'])
        
        # 檢查是否有 U(V), I(A), P(W), WP(Wh) 欄位，若無則補上並填入 0
        required_columns = ['U(V)', 'I(A)', 'P(W)', 'WP(Wh)']
        for col in required_columns:
            if col not in df.columns:
                df[col] = 0

        # 將 start_date, start_time, end_date, end_time 轉換為 datetime 格式
        start = pd.to_datetime(f"{start_date.get()} {start_time.get()}")
        end = pd.to_datetime(f"{end_date.get()} {end_time.get()}")
        # 計算 start 和 end 之間的分鐘數
        minutes_difference = int((end - start).total_seconds() / 60)
        
        # 過濾指定的日期時間範圍，並創建副本
        filtered_df = df[(df['datetime'] >= start) & (df['datetime'] <= end)].copy()

        if filtered_df.empty:
            messagebox.showinfo("結果", "指定範圍內沒有資料！")
            return

        # 計算平均值
        averages = filtered_df.mean(numeric_only=True)

        # 計算電力啟停周期
        power_column = 'P(W)'
        if power_column in filtered_df.columns:
            # 確保正確建立 power_on 欄位
            filtered_df.loc[:, 'power_on'] = filtered_df[power_column] >= 3

            # 計算啟停周期次數
            power_cycles = int(filtered_df['power_on'].astype(int).diff().fillna(0).abs().sum() // 2)

            # 計算大於等於3W和小於3W的週期數，排除頭尾兩個周期
            mask = filtered_df['power_on']
            groups = (mask != mask.shift()).cumsum()
            segments = pd.DataFrame({
                '狀態': mask,
                '區段編號': groups,
                '時間': filtered_df['datetime']
            }).groupby(['區段編號', '狀態']).agg({'時間': ['min', 'max']}).reset_index()

            # 排除頭尾兩個周期
            if len(segments) > 2:
                segments = segments.iloc[1:-1]

            # 計算每個區段的持續時間
            segments['持續時間'] = (segments[('時間', 'max')] - segments[('時間', 'min')]).dt.total_seconds()

            # 分別計算大於等於3W和小於3W的週期數與平均時間
            above_segments = segments[segments['狀態']]
            below_segments = segments[~segments['狀態']]

            above_count = len(above_segments)
            below_count = len(below_segments)

            above_avg_time = (above_segments['持續時間'].mean() / 60) if above_count > 0 else 0
            below_avg_time = (below_segments['持續時間'].mean() / 60) if below_count > 0 else 0

            # 計算百分比
            if above_avg_time + below_avg_time > 0:
                above_percentage = (above_avg_time / (above_avg_time + below_avg_time)) * 100
            else:
                above_percentage = 0
        else:
            power_cycles = "無法計算，缺少 P(W) 欄位"
            
        # 計算 WP(Wh) 欄位的差值
        wp_column = 'WP(Wh)'
        if wp_column in filtered_df.columns:
            wp_difference = filtered_df[wp_column].iloc[-1] - filtered_df[wp_column].iloc[0]
            
            # 使用線性法推算 24 小時的差值
            total_seconds = (filtered_df['datetime'].iloc[-1] - filtered_df['datetime'].iloc[0]).total_seconds()
            if (total_seconds > 0):
                wp_24h_difference = (wp_difference / total_seconds) * (24 * 3600)
            else:
                wp_24h_difference = "無法計算，時間範圍不足"
        else:
            wp_difference = "無法計算，缺少 WP(Wh) 欄位"
            wp_24h_difference = "無法計算，缺少 WP(Wh) 欄位"

        # 顯示結果
        result_textbox.delete(1.0, tk.END)  # 清空文字框
        result_textbox.insert(tk.END, f"統計範圍：{start} ~ {end}\n")
        result_textbox.insert(tk.END, "平均值計算：\n")
        for column, avg in averages.items():
            result_textbox.insert(tk.END, f"{column}: {avg:.2f}\n")
        result_textbox.insert(tk.END, f"\nON / Off 周期次數：{power_cycles}\n")
        result_textbox.insert(tk.END, f"On 的平均時間: {above_avg_time:.1f} 分\n" if above_count > 0 else "P(W) >= 3 的平均時間: 無資料\n")
        result_textbox.insert(tk.END, f"Off 的平均時間: {below_avg_time:.1f} 分\n" if below_count > 0 else "P(W) < 3 的平均時間: 無資料\n")
        result_textbox.insert(tk.END, f"On / Off 百分比: {above_percentage:.2f}%\n")
        result_textbox.insert(tk.END, f"\n電力消耗：{wp_difference:.2f} w / {minutes_difference} 分\n")
        result_textbox.insert(tk.END, f"24 小時電力消耗：{wp_24h_difference:.1f} w\n")
    except Exception as e:
        messagebox.showerror("錯誤", f"計算平均值或電力啟停周期時發生錯誤：{e}")

# 新增儲存結果的函數
def save_results():
    try:
        # 選擇儲存檔案的路徑
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if not file_path:
            return  # 如果使用者取消操作，直接返回

        # 取得多行文字框的內容
        content = result_textbox.get(1.0, tk.END).strip()

        # 將內容寫入檔案
        with open(file_path, "w", encoding="utf-8") as file:
            file.write(content)

        messagebox.showinfo("成功", "結果已成功儲存！")
    except Exception as e:
        messagebox.showerror("錯誤", f"儲存結果時發生錯誤：{e}")

# GUI 元件
tk.Label(root, text="CSV 檔案路徑:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
tk.Entry(root, textvariable=csv_path, width=50).grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="選擇檔案", command=select_file).grid(row=0, column=2, padx=5, pady=5)

tk.Label(root, text="圖表標題:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
tk.Entry(root, textvariable=chart_title, width=50).grid(row=1, column=1, padx=5, pady=5)

tk.Button(root, text="繪製圖表", command=lambda: plot_chart()).grid(row=4, column=0, pady=10)
tk.Button(root, text="計算平均值", command=calculate_statistics).grid(row=4, column=1, pady=10)
tk.Button(root, text="儲存結果", command=save_results).grid(row=4, column=2, pady=10)

# 多行文字框顯示結果
result_frame = tk.Frame(root)  # 使用 Frame 包含文字框
result_frame.grid(row=5, column=0, columnspan=3,rowspan=4, padx=5, pady=10)
# 
# 開始日期與時間
tk.Label(root, text="資料日期:").grid(row=4, column=3, columnspan=2, padx=5, pady=5, sticky="nsew")
def increment_date_time(var, increment, unit):
    try:
        current_value = pd.to_datetime(var.get())
        if unit == "day":
            new_value = current_value + pd.Timedelta(days=increment)
        elif unit == "hour":
            new_value = current_value + pd.Timedelta(hours=increment)
        var.set(new_value.strftime('%Y-%m-%d' if unit == "day" else '%H:%M'))
    except Exception:
        messagebox.showerror("錯誤", "無效的日期或時間格式！")

def bind_increment(widget, var, unit):
    def on_key(event):
        if event.state & 0x4:  # 檢查是否按下 CTRL 鍵
            if event.keysym == "Up":
                increment_date_time(var, 1, unit)
            elif event.keysym == "Down":
                increment_date_time(var, -1, unit)
    widget.bind("<KeyPress-Up>", on_key)
    widget.bind("<KeyPress-Down>", on_key)


start_date_entry = tk.Entry(root, textvariable=start_date, width=10)
start_date_entry.grid(row=5, column=3, padx=5, pady=5, sticky="w")
bind_increment(start_date_entry, start_date, "day")

start_time_entry = tk.Entry(root, textvariable=start_time, width=10)
start_time_entry.grid(row=6, column=3, padx=5, pady=5, sticky="w")
bind_increment(start_time_entry, start_time, "hour")

end_date_entry = tk.Entry(root, textvariable=end_date, width=10)
end_date_entry.grid(row=7, column=3, padx=5, pady=5, sticky="w")
bind_increment(end_date_entry, end_date, "day")

end_time_entry = tk.Entry(root, textvariable=end_time, width=10)
end_time_entry.grid(row=8, column=3, padx=5, pady=5, sticky="w")
bind_increment(end_time_entry, end_time, "hour")

# 創建多行文字框
result_textbox = tk.Text(
    result_frame, 
    width=70, 
    height=15, 
    wrap=tk.WORD  # 自動換行
)

result_textbox.pack(fill="both", expand=True)  # 文字框填滿剩餘空間

# 啟動主迴圈
root.mainloop()