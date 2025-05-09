#data log plot and EEF simple calculate for VM7000/PW3335 Data Collection
#-------------------------------------------------------------------------------
#Rev.1.4 新增日期時間增減功能
#Rev.2.0 新增光棒移動來設定計算範圍,將圖表改為內崁式
#-------------------------------------------------------------------------------
# 
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.dates import DateFormatter
from matplotlib import rcParams
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
import sys  # 引入 sys 模組以便退出程式

# 在程式中的適當位置（例如在plot_chart函數之前）添加DraggableLine類
class DraggableLine:
    def __init__(self, ax, xdata, ydata, initial_pos, color='red', linestyle='--', linewidth=1, 
                 date_var=None, time_var=None):
        self.ax = ax
        self.xdata = xdata
        self.ydata = ydata
        self.line = ax.axvline(x=initial_pos, color=color, linestyle=linestyle, linewidth=linewidth)
        self.press = None
        self.date_var = date_var
        self.time_var = time_var
        self.cid_press = self.line.figure.canvas.mpl_connect('button_press_event', self.on_press)
        self.cid_release = self.line.figure.canvas.mpl_connect('button_release_event', self.on_release)
        self.cid_motion = self.line.figure.canvas.mpl_connect('motion_notify_event', self.on_motion)
        
    def on_press(self, event):
        if event.inaxes != self.ax:
            return
        contains, attrd = self.line.contains(event)
        if not contains:
            return
        self.press = True
        
    def on_motion(self, event):
        if not self.press or event.inaxes != self.ax:
            return
        x_pos = event.xdata
        self.line.set_xdata([x_pos, x_pos])
        self.update_text_boxes(x_pos)
        self.line.figure.canvas.draw()
        
    def on_release(self, event):
        self.press = False
        self.line.figure.canvas.draw()
        
    def get_position(self):
        return self.line.get_xdata()[0]
        
    def update_text_boxes(self, x_pos):
        if self.date_var is not None and self.time_var is not None:
            dt = mdates.num2date(x_pos)
            self.date_var.set(dt.strftime('%Y-%m-%d'))
            self.time_var.set(dt.strftime('%H:%M'))

# 修改plot_chart函數，在繪製圖表後添加光棒
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

        # 確認電力欄位是否存在
        power_columns = ['U(V)', 'I(A)', 'P(W)']
        if all(col in df.columns for col in power_columns):
            df_power = df[['datetime'] + power_columns]
        else:
            messagebox.showerror("錯誤", "缺少電力欄位！")
            return
        
        # 其餘欄位視為溫度資料
        temp_columns = [col for col in df.columns if col not in ['Date', 'Time', 'datetime', 'WP(Wh)'] + power_columns]
        df_temp = df[['datetime'] + temp_columns]
        
        # 清除舊圖表
        ax1.clear()
        ax2.clear()

        # 繪製圖表
        fig.suptitle(chart_title.get())
        
        # 繪製溫度資料
        for col in temp_columns:
            ax1.plot(df_temp['datetime'], df_temp[col], label=col)
        ax1.set_ylabel("溫度")
        #ax1.legend()
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
        ax2.xaxis.set_major_locator(plt.MaxNLocator(10))  # 限制 X 軸最多顯示 10 個標籤
        plt.setp(ax2.xaxis.get_majorticklabels(), rotation=45)
        plt.tight_layout()

        # 添加可拖動的光棒
        global start_line, end_line  # 聲明為全域變數以便在其他函數中訪問
        
        # 獲取開始和結束時間
        start_pos = pd.to_datetime(f"{start_date.get()} {start_time.get()}")
        end_pos = pd.to_datetime(f"{end_date.get()} {end_time.get()}")
        
        # 創建可拖動的光棒，並傳遞對應的文字框變數
        start_line = DraggableLine(ax1, df_temp['datetime'], df_temp[temp_columns[0]], 
                                 start_pos, color='blue', linestyle='--',
                                 date_var=start_date, time_var=start_time)
        end_line = DraggableLine(ax1, df_temp['datetime'], df_temp[temp_columns[0]], 
                               end_pos, color='red', linestyle='--',
                               date_var=end_date, time_var=end_time)
        
        # 更新嵌入的圖表
        canvas.draw()

    except Exception as e:
        messagebox.showerror("錯誤", f"繪製圖表時發生錯誤：{e}")

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

# 定義視窗關閉事件處理函數
def on_closing():
    if messagebox.askokcancel("退出", "確定要退出程式嗎？"):
        root.destroy()  # 關閉 Tkinter 視窗
        sys.exit()  # 確保程式完全退出


# 設定 matplotlib 使用的字體
rcParams['font.sans-serif'] = ['Microsoft JhengHei']  # 使用微軟正黑體
rcParams['axes.unicode_minus'] = False  # 解決負號無法顯示的問題

# 初始化主視窗
root = tk.Tk()
root.title("CSV Plotter with Statistics")

# 全域變數
csv_path = tk.StringVar()
chart_title = tk.StringVar()
start_datetime = tk.StringVar()
end_datetime = tk.StringVar()
start_date = tk.StringVar()
start_time = tk.StringVar()
end_date = tk.StringVar()
end_time = tk.StringVar()

# GUI 元件
main_frame = tk.Frame(root)  # 使用 Frame 包含文字框
main_frame.grid(row=0, column=0, padx=5, pady=10)
tk.Label(main_frame, text="CSV 檔案路徑:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
tk.Entry(main_frame, textvariable=csv_path, width=50).grid(row=0, column=1, padx=5, pady=5)
tk.Button(main_frame, text="選擇檔案", command=select_file).grid(row=0, column=2, padx=5, pady=5)

tk.Label(main_frame, text="圖表標題:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
tk.Entry(main_frame, textvariable=chart_title, width=50).grid(row=1, column=1, padx=5, pady=5)

tk.Button(main_frame, text="繪製圖表", command=lambda: plot_chart()).grid(row=4, column=0, pady=10)
tk.Button(main_frame, text="計算平均值", command=calculate_statistics).grid(row=4, column=1, pady=10)
tk.Button(main_frame, text="儲存結果", command=save_results).grid(row=4, column=2, pady=10)

# 多行文字框顯示結果
result_frame = tk.Frame(root)  # 使用 Frame 包含文字框
result_frame.grid(row=1, column=0, padx=5, pady=10)
# 
# 創建多行文字框
result_textbox = tk.Text(
    result_frame, 
    width=70, 
    height=30, 
    wrap=tk.WORD  # 自動換行
)
result_textbox.grid(row=0, rowspan=5, column=0, padx=5, pady=5)

# 開始日期與時間
tk.Label(result_frame, text="資料日期:").grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky="nsew")
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


start_date_entry = tk.Entry(result_frame, textvariable=start_date, width=10)
start_date_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
bind_increment(start_date_entry, start_date, "day")

start_time_entry = tk.Entry(result_frame, textvariable=start_time, width=10)
start_time_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
bind_increment(start_time_entry, start_time, "hour")

end_date_entry = tk.Entry(result_frame, textvariable=end_date, width=10)
end_date_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")
bind_increment(end_date_entry, end_date, "day")

end_time_entry = tk.Entry(result_frame, textvariable=end_time, width=10)
end_time_entry.grid(row=4, column=1, padx=5, pady=5, sticky="w")
bind_increment(end_time_entry, end_time, "hour")

fig, (ax1, ax2) = plt.subplots(2, 1, sharex=True, figsize=(8, 6), facecolor='lightgray')
canvas = FigureCanvasTkAgg(fig, master=root)
canvas_widget = canvas.get_tk_widget()
canvas_widget.grid(row=0, rowspan=2, column=1, padx=5, pady=5)

# 添加 Matplotlib 的工具列
toolbar_frame = tk.Frame(root)
toolbar_frame.grid(row=2, column=1, padx=5, pady=5, sticky="nsew")
toolbar = NavigationToolbar2Tk(canvas, toolbar_frame)
toolbar.update()

# 綁定視窗關閉事件
root.protocol("WM_DELETE_WINDOW", on_closing)

# 啟動主迴圈
root.mainloop()
