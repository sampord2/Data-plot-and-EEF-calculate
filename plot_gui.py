#data log plot and EEF simple calculate for VM7000/PW3335 Data Collection
#-------------------------------------------------------------------------------
#Rev.1.4 新增日期時間增減功能
#Rev.2.0 新增光棒移動來設定計算範圍,將圖表改為內崁式,新增能耗計算模組
#Rev.2.1 繪製新數據前，確保所有相關軸都被正確清除和重置
#Rev.2.2 修正語法錯誤,圖表比例改為7:3
#-------------------------------------------------------------------------------
# 
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import MaxNLocator
from matplotlib import rcParams
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends._backend_tk import NavigationToolbar2Tk
import sys  # 引入 sys 模組以便退出程式
import os

def resource_path(relative_path):
    """ 獲取資源的絕對路徑 """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(getattr(sys, '_MEIPASS'), relative_path)
    return os.path.join(os.path.abspath('.'), relative_path)

class EnergyCalculator:
    def __init__(self):
        pass
    def current_ef_thresholds(self,energy_allowance,fridge_type):
        if fridge_type == 5:
            #IF(fridge_type=5,ROUND(N4*1.72,1),ROUND(N4*1.6,1))
            threshold_lv1 = round(energy_allowance * 1.72,1)
            threshold_lv2 = round(energy_allowance * 1.54,1)
            threshold_lv3 = round(energy_allowance * 1.36,1)
            threshold_lv4 = round(energy_allowance * 1.18,1)
        else:
            #IF(fridge_type=5,ROUND(N4*1.72,1),ROUND(N4*1.6,1))
            threshold_lv1 = round(energy_allowance * 1.6,1)
            threshold_lv2 = round(energy_allowance * 1.45,1)
            threshold_lv3 = round(energy_allowance * 1.3,1)
            threshold_lv4 = round(energy_allowance * 1.15,1)
        return[ threshold_lv1, threshold_lv2, threshold_lv3, threshold_lv4 ]

    def future_ef_thresholds(self,energy_allowance,fridge_type):
        if fridge_type == 5:
            #IF(fridge_type=5,ROUND(N4*1.72,1),ROUND(N4*1.6,1))
            threshold_lv1 = round(energy_allowance * 1.294,1)
            threshold_lv2 = round(energy_allowance * 1.221,1)
            threshold_lv3 = round(energy_allowance * 1.147,1)
            threshold_lv4 = round(energy_allowance * 1.074,1)
        else:
            #IF(fridge_type=5,ROUND(N4*1.72,1),ROUND(N4*1.6,1))
            threshold_lv1 = round(energy_allowance * 1.308,1)
            threshold_lv2 = round(energy_allowance * 1.231,1)
            threshold_lv3 = round(energy_allowance * 1.154,1)
            threshold_lv4 = round(energy_allowance * 1.077,1)
        return[ threshold_lv1, threshold_lv2, threshold_lv3, threshold_lv4 ]


    def calculate(self, VF, VR, daily_consumption, fridge_temp, freezer_temp, fan_type):
        """
        計算冰箱能耗相關指標
        
        參數:
            VR: 冷藏室容積(L)
            VF: 冷凍室容積(L)
            daily_consumption: 日耗電量(kWh/日)
            fridge_temp: 冷藏室溫度(°C), 預設3.0
            freezer_temp: 冷凍室溫度(°C), 預設-18.0
        
        返回:
            包含所有計算結果的字典
        """
        results = {}
        
        # 1. 計算K值 (溫度係數)
        K = self.calculate_K_value(freezer_temp, fridge_temp)
        #print(f"K值: {K}")
        # 2. 計算等效內容積
        equivalent_volume = self.calculate_equivalent_volume(VR, VF, K)
        
        # 3. 確定冰箱型式
        fridge_type = self.determine_fridge_type(equivalent_volume, VR, VF, fan_type)
        #print(f"冰箱型式: {fridge_type}")
        # 4. 計算容許耗用能源基準 (每月)
        energy_allowance = self.calculate_energy_allowance(equivalent_volume, fridge_type)
        
        # 5. 計算2027容許耗用能源基準
        future_energy_allowance = self.calculate_future_energy_allowance(equivalent_volume, fridge_type)
        
        # 6. 計算耗電量基準 (每月)
        benchmark_consumption = self.calculate_benchmark_consumption(equivalent_volume, energy_allowance)
        
        # 7. 計算2027耗電量基準
        future_benchmark_consumption = self.calculate_future_benchmark_consumption(equivalent_volume, future_energy_allowance)
        
        # 8. 計算實測月耗電量
        monthly_consumption = round(daily_consumption * 30,1)
        
        # 9. 計算EF值 (能效因子)
        if monthly_consumption == 0:
            ef_value = 0.0
        else:
            ef_value = round(equivalent_volume / monthly_consumption,1)
        
        # 9.1 計算現有效率基準百分比和等級
        current_ef_thresholds = self.current_ef_thresholds(energy_allowance, fridge_type)

        # 10. 計算現有效率基準百分比和等級
        current_percent, current_grade = self.calculate_current_efficiency(ef_value, current_ef_thresholds)
        
        # 10.1 計算2027新效率基準百分比和等級
        future_ef_thresholds = self.future_ef_thresholds(future_energy_allowance, fridge_type)

        # 11. 計算2027新效率基準百分比和等級
        future_percent, future_grade = self.calculate_future_efficiency(ef_value, future_ef_thresholds)
        
        # 整理所有結果
        results.update({
            '冷凍室溫度': freezer_temp,
            '冷藏室溫度': fridge_temp,
            'K值': K,
            'VF(L)': VF,
            'VR(L)': VR,
            '等效內容積(L)': equivalent_volume,
            '冰箱型式': fridge_type,
            '\n----能效相關計算結果----': '',
            'EF值': ef_value,
            '實測月耗電量(kWh/月)': monthly_consumption,
            '2018年容許耗用能源基準(L/kWh/月)': energy_allowance,
            '2018年耗電量基準(kWh/月)': benchmark_consumption,
            '2018年一級效率EF值': current_ef_thresholds[0],
            '2018年效率等級': current_grade,
            '2018年一級效率百分比(%)': current_percent,
            '\n----2027年新能效公式----': '',
            '2027容許耗用能源基準(L/kWh/月)': future_energy_allowance,
            '2027年耗電量基準(kWh/月)': future_benchmark_consumption,
            '2027年一級效率EF值': future_ef_thresholds[0],
            '2027年效率等級': future_grade,
            '2027年一級效率百分比(%)': future_percent
        })
        #print(f"threshold: {current_ef_thresholds}, {future_ef_thresholds}")
        return results
    
    def calculate_K_value(self, freezer_temp, fridge_temp):
        """計算K值 (溫度係數)"""
        # 根據公式 K = (30 - 冷凍庫溫度) / (30 - 冷藏庫溫度)
        #print(f"冷凍庫溫度: {freezer_temp}, 冷藏庫溫度: {fridge_temp}")        
        return round((30 - freezer_temp) / (30 - fridge_temp), 2)
    
    def calculate_equivalent_volume(self, VR, VF, K):
        """計算等效內容積"""
        return round(VR + (K * VF), 1)
    
    def determine_fridge_type(self, equivalent_volume, VR, VF, fan_type):
        """確定冰箱型式"""
        if VF == 0:  # 只有冷藏室
            return 5
        elif equivalent_volume < 400 and fan_type == 1:
            return 1  # 假設是風冷式(實際應根據具體設計)
        elif equivalent_volume >= 400 and fan_type == 1:
            return 2
        elif equivalent_volume < 400 and fan_type == 0:
            return 3
        else:
            return 4  # 假設是風冷式(實際應根據具體設計)
    
    def calculate_energy_allowance(self, equivalent_volume, fridge_type):
        """計算容許耗用能源基準"""
        # 根據公式，ROUND(IFS(fridge_type=1,equivalent_volume/(0.037*equivalent_volume+24.3),fridge_type=2,equivalent_volume/(0.031*M4+21),fridge_type=3,equivalent_volume/(0.033*equivalent_volume+19.7),fridge_type=4,equivalent_volume/(0.029*equivalent_volume+17),fridge_type=5,equivalent_volume/(0.033*equivalent_volume+15.8)),1)
        if fridge_type == 1:
            return round( equivalent_volume / (0.037 * equivalent_volume + 24.3), 1)
        elif fridge_type == 2:
            return round( equivalent_volume / (0.031 * equivalent_volume + 21), 1)
        elif fridge_type == 3:
            return round( equivalent_volume / (0.033 * equivalent_volume + 19.7), 1)
        elif fridge_type == 4:
            return round( equivalent_volume / (0.029 * equivalent_volume + 17), 1)
        else:
            return round( equivalent_volume / (0.033 * equivalent_volume + 15.8), 1)
    
    def calculate_future_energy_allowance(self, equivalent_volume, fridge_type):
        """計算2027年容許耗用能源基準"""
        # 公式:=ROUND(IFS(F4=1,1.3*M4/(0.037*M4+24.3),F4=2,1.3*M4/(0.031*M4+21),F4=3,1.3*M4/(0.033*M4+19.7),F4=4,1.3*M4/(0.029*M4+17),F4=5,1.36*M4/(0.033*M4+15.8)),1)
        # F4 = fridge_type, M4 = equivalent_volume
        if fridge_type == 1:
            return round( 1.3 * equivalent_volume / (0.037 * equivalent_volume + 24.3), 1)
        elif fridge_type == 2:
            return round( 1.3 * equivalent_volume / (0.031 * equivalent_volume + 21), 1)
        elif fridge_type == 3:
            return round(1.3 * equivalent_volume / (0.033 * equivalent_volume + 19.7), 1)
        elif fridge_type == 4:
            return round(1.3 * equivalent_volume / (0.029 * equivalent_volume + 17), 1)
        else:
            return round(1.36 * equivalent_volume / (0.033 * equivalent_volume + 15.8), 1)

    def calculate_benchmark_consumption(self, equivalent_volume, energy_allowance):
        """計算耗電量基準"""
        # 根據公式:ROUND(equivalent_volume / energy_allowance, 1)
        return round(equivalent_volume / energy_allowance, 1)
    
    def calculate_future_benchmark_consumption(self, equivalent_volume, future_energy_allowance):
        """計算2027耗電量基準"""
        return round(equivalent_volume / future_energy_allowance, 1)
    
    def calculate_current_efficiency(self, ef_value, thresholds):
        # 確定等級
        if ef_value >= thresholds[0]:
            grade = "1級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[0] * 0.95:
            grade = "1*級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[1]:
            grade = "2級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[2]:
            grade = "3級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[3]:
            grade = "4級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        else :
            grade = "5級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        
        return final_percent, grade
    
    def calculate_future_efficiency(self, ef_value, thresholds):
        # 確定等級
        if ef_value >= thresholds[0]:
            grade = "1級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[0] * 0.95:
            grade = "1*級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[1]:
            grade = "2級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[2]:
            grade = "3級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[3]:
            grade = "4級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        else :
            grade = "5級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        
        return final_percent, grade


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
            self.time_var.set(dt.strftime('%H:%M:%S'))

def plot_chart():
    ax1_color = "lightcyan"
    ax2_color = "lightyellow"
    try:
        global ax1, ax2, ax2a, ax2b, ax2c
        
        # 清除現有繪圖
        fig.clf()

        # 重新建立 gridspec 與子圖，保留 7:3 高度比例
        gs = fig.add_gridspec(2, 1, height_ratios=[7, 3])
        ax1 = fig.add_subplot(gs[0, 0], facecolor=ax1_color)
        ax2 = fig.add_subplot(gs[1, 0], sharex=ax1, facecolor=ax2_color)
        
        # 清除右側的雙 Y 軸
        for ax in fig.axes:
            if ax not in [ax1, ax2]:
                ax.remove()
        
        # 重新創建右側 Y 軸
        ax2a = ax2  # 左側 Y 軸 (功率 P(W))
        ax2b = ax2.twinx()  # 右側第一個 Y 軸 (電壓 U(V))
        ax2c = ax2.twinx()  # 右側第二個 Y 軸 (電流 I(A))
        
        # 調整右側第二個 Y 軸的位置
        ax2c.spines['right'].set_position(('outward', 35))

        # 讀取 CSV 檔案（指定編碼以支援中文欄位）
        try:
            df = pd.read_csv(csv_path.get(), encoding="utf-8-sig")
        except UnicodeDecodeError:
            df = pd.read_csv(csv_path.get(), encoding="big5")
        # 合併 Date 和 Time 欄位為 datetime
        df['datetime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'])

        # 去除第一筆資料為空白的欄位
        df = df.dropna(how='all', axis=1)
        
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
        temp_columns = [col for col in df.columns if col not in ['Date', 'Time', 'datetime'] + required_columns]
        df_temp = df[['datetime'] + temp_columns]
        
        # 繪製圖表
        fig.suptitle(chart_title.get())
        
        # 繪製溫度資料
        for col in temp_columns:
            ax1.plot(df_temp['datetime'], df_temp[col], label=col)
        ax1.set_ylabel("溫度")

        ax1.legend(temp_columns, loc='upper left')
        ax1.set_facecolor(ax1_color)
        ax1.grid(True, linestyle='dashdot')
        ax1.tick_params(axis='x', which='both', bottom=False, top=False, labelbottom=False)
        
        # 繪製電力資料
        ax2a.plot(df_power['datetime'], df_power['P(W)'], label='P(W)', color='green')
        ax2a.set_ylabel("功率 (P)", color='green')
        ax2a.tick_params(axis='y', labelcolor='green')
        ax2a.grid(True, linestyle='dashdot')

        ax2b.plot(df_power['datetime'], df_power['U(V)'], label='U(V)', color='blue')
        ax2b.set_ylabel("電壓 (U)", color='blue')
        ax2b.tick_params(axis='y', labelcolor='blue')
        ax2b.set_ylim(0, 120)

        ax2c.plot(df_power['datetime'], df_power['I(A)'], label='I(A)', color='red')
        ax2c.set_ylabel("電流 (I)", color='red')
        ax2c.tick_params(axis='y', labelcolor='red')

        ax2.set_facecolor(ax2_color)
        ax2.xaxis.set_major_locator(MaxNLocator(10))
        plt.setp(ax2.xaxis.get_majorticklabels(), rotation=45)
        plt.tight_layout()
        # 設定 ax2 的 x 軸 datetime 顯示格式為 "d-hh:mm"
        ax2.xaxis.set_major_formatter(mdates.DateFormatter('%d-%H:%M'))
        ax2.set_xlim(ax1.get_xlim())

        # 添加可拖動的光棒
        global start_line, end_line
        start_pos = pd.to_datetime(f"{start_date.get()} {start_time.get()}")
        end_pos = pd.to_datetime(f"{end_date.get()} {end_time.get()}")
        
        start_line = DraggableLine(ax1, df_temp['datetime'], df_temp[temp_columns[0]], 
                                 start_pos, color='blue', linestyle='--',
                                 date_var=start_date, time_var=start_time)
        end_line = DraggableLine(ax1, df_temp['datetime'], df_temp[temp_columns[0]], 
                               end_pos, color='red', linestyle='--',
                               date_var=end_date, time_var=end_time)
        
        canvas.draw()
        toolbar.update()

    except Exception as e:
        messagebox.showerror("錯誤", f"繪製圖表時發生錯誤：{e}")

def calculate_statistics():
    try:
        # 讀取 CSV 檔案
        try:
            df = pd.read_csv(csv_path.get(), encoding="utf-8-sig")
        except UnicodeDecodeError:
            df = pd.read_csv(csv_path.get(), encoding="big5")
        
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

            
        # 計算能耗
        fan_type = fan_type_var.get()  # 取得風扇類型的狀態
        vf = float(vf_entry_var.get()) if vf_entry_var.get().isdigit() else 0
        vr = float(vr_entry_var.get()) if vr_entry_var.get().isdigit() else 0
        fridge_temp = float(temp_r_entry_var.get()) if temp_r_entry_var.get().replace('.', '', 1).isdigit() else 3  # 冷藏室溫度
        freezer_temp = float(temp_f_entry_var.get()) if temp_f_entry_var.get().replace('.', '', 1).lstrip('-').isdigit() else -18.0  # 冷凍室溫度
        energy_calculator = EnergyCalculator()
        if isinstance(wp_24h_difference, (int, float)) and vf > 0 and vr > 0:
            daily_consumption = wp_24h_difference / 1000  # 將 Wh 轉換為 kWh
            # 計算
            results = energy_calculator.calculate(vf, vr, daily_consumption, fridge_temp, freezer_temp, fan_type)
            # 提取結果
            #if results:
                # 打印結果
                #print("冰箱能耗計算結果:")
                #for key, value in results.items():
                    #print(f"{key}: {value}")
        else:
            results = None
            print("無耗電量數據,無法計算能耗")

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
        result_textbox.insert(tk.END, f"\n能耗計算：\n")
        if results:
            for key, value in results.items():
                result_textbox.insert(tk.END, f"{key}: {value}\n")
        else:
            result_textbox.insert(tk.END, "無法計算能耗，請檢查數據\n")
    except Exception as e:
        messagebox.showerror("錯誤", f"計算平均值或電力啟停周期時發生錯誤：{e}")
        print(f"caculate error：{e}")

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
            try:
                df = pd.read_csv(file_path, encoding="utf-8-sig")
            except UnicodeDecodeError:
                df = pd.read_csv(file_path, encoding="big5")
            # 檢查是否有 Date 和 Time 欄位
            if 'Date' in df.columns and 'Time' in df.columns:
                # 合併 Date 和 Time 欄位為 datetime
                df['datetime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'])
                
                # 設定開始和結束日期時間
                start_date.set(df['datetime'].iloc[0].strftime('%Y-%m-%d'))
                start_time.set(df['datetime'].iloc[0].strftime('%H:%M:%S'))
                end_date.set(df['datetime'].iloc[-1].strftime('%Y-%m-%d'))
                end_time.set(df['datetime'].iloc[-1].strftime('%H:%M:%S'))
            else:
                messagebox.showerror("錯誤", "CSV 檔案缺少 Date 或 Time 欄位！")

            # 檢查是否有 U(V), I(A), P(W), WP(Wh) 欄位，若無則補上並填入 0
            required_columns = ['U(V)', 'I(A)', 'P(W)', 'WP(Wh)']
            for col in required_columns:
                if col not in df.columns:
                    df[col] = 0

        except Exception as e:
            messagebox.showerror("錯誤", f"處理 CSV 檔案時發生錯誤：{e}")
            print(f"Select_file: err：{e}")

# 定義視窗關閉事件處理函數
def on_closing():
    if messagebox.askokcancel("退出", "確定要退出程式嗎？"):
        root.destroy()  # 關閉 Tkinter 視窗
        sys.exit()  # 確保程式完全退出


# 設定 matplotlib 使用的字體
rcParams['font.sans-serif'] = ['Microsoft JhengHei']  # 使用微軟正黑體
rcParams['axes.unicode_minus'] = False  # 解決負號無法顯示的問題
rcParams['font.size'] = 8
rcParams['axes.titlesize'] = 8      # 座標軸標題字體大小
rcParams['axes.labelsize'] = 8      # 座標軸標籤字體大小
rcParams['xtick.labelsize'] = 8     # X軸刻度字體大小
rcParams['ytick.labelsize'] = 8     # Y軸刻度字體大小
rcParams['legend.fontsize'] = 8     # 圖例字體大小
rcParams['figure.titlesize'] = 12    # 圖形標題字體大小

# 初始化主視窗
root = tk.Tk()
root.title("CSV Plotter with Statistics 2.2")
root.iconbitmap(resource_path('favicon.ico'))

# 全域變數
csv_path = tk.StringVar()
chart_title = tk.StringVar()
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
    width=50, 
    height=30, 
    wrap=tk.WORD  # 自動換行
)
result_textbox.grid(row=0, rowspan=12, column=0, padx=5, pady=5)

# 能耗計算用欄位
tk.Label(result_frame, text="能耗計算用:").grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky="nsew")
tk.Label(result_frame, text="冷凍室容積(L):").grid(row=1, column=1, padx=5, pady=5, sticky="w")
vf_entry_var = tk.StringVar(value="150")
vf_entry = tk.Entry(result_frame, width=5, textvariable=vf_entry_var)
vf_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
tk.Label(result_frame, text="F:").grid(row=1, column=2, padx=5, pady=5, sticky="w")
temp_f_entry_var = tk.StringVar(value="-18.0")
temp_f_entry = tk.Entry(result_frame, width=5, textvariable=temp_f_entry_var)
temp_f_entry.grid(row=2, column=2, padx=5, pady=5, sticky="w")
tk.Label(result_frame, text="冷藏室容積(L):").grid(row=3, column=1, padx=5, pady=5, sticky="w")
vr_entry_var = tk.StringVar(value="350")
vr_entry = tk.Entry(result_frame, width=5, textvariable=vr_entry_var)
vr_entry.grid(row=4, column=1, padx=5, pady=5, sticky="w")
tk.Label(result_frame, text="R:").grid(row=3, column=2, padx=5, pady=5, sticky="w")
temp_r_entry_var = tk.StringVar(value="3.0")
temp_r_entry = tk.Entry(result_frame, width=5, textvariable=temp_r_entry_var)
temp_r_entry.grid(row=4, column=2, padx=5, pady=5, sticky="w")
fan_type_var = tk.IntVar(value=1)  # 0: unchecked, 1: checked
fan_type_checkbox = tk.Checkbutton(result_frame, text="風扇式", variable=fan_type_var)
fan_type_checkbox.grid(row=5, column=1, padx=5, pady=5, sticky="w")

# 分割線
ttk.Separator(result_frame, orient="horizontal").grid(row=6, column=1, sticky="ew", pady=10)
# 開始日期與時間
tk.Label(result_frame, text="計算範圍:").grid(row=7, column=1, columnspan=2, padx=5, pady=5, sticky="nsew")
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
start_date_entry.grid(row=8, column=1, padx=5, pady=5, sticky="w")
bind_increment(start_date_entry, start_date, "day")

start_time_entry = tk.Entry(result_frame, textvariable=start_time, width=10)
start_time_entry.grid(row=9, column=1, padx=5, pady=5, sticky="w")
bind_increment(start_time_entry, start_time, "hour")

end_date_entry = tk.Entry(result_frame, textvariable=end_date, width=10)
end_date_entry.grid(row=10, column=1, padx=5, pady=5, sticky="w")
bind_increment(end_date_entry, end_date, "day")

end_time_entry = tk.Entry(result_frame, textvariable=end_time, width=10)
end_time_entry.grid(row=11, column=1, padx=5, pady=5, sticky="w")
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