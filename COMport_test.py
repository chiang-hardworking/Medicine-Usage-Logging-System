import customtkinter as ctk
import tkinter.ttk as ttk
import tkinter.messagebox as messagebox
import serial
import serial.tools.list_ports
import threading
import datetime
import os
import json
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

# 設定主題
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

CONFIG_FILE = "config.json"

class BalanceDebugGUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("天平數據偵錯系統 - 原始數據監控中")
        self.geometry("1000x800")

        # --- 載入設定 ---
        self.config = self.load_config()
        self.file_name = "medicine_data.xlsx"
        self.is_monitoring = False
        self.ser = None

        # --- UI 配置 (維持原有佈局) ---
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # 左側控制面板
        self.sidebar = ctk.CTkFrame(self, width=240, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        ctk.CTkLabel(self.sidebar, text="系統偵錯設定", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=15)

        # Port 選擇
        ctk.CTkLabel(self.sidebar, text="選擇 COM Port:").pack(pady=(5, 0))
        available_ports = self.get_ports()
        self.port_menu = ctk.CTkOptionMenu(self.sidebar, values=available_ports)
        self.port_menu.pack(pady=5, padx=20)
        if self.config.get("last_port") in available_ports:
            self.port_menu.set(self.config.get("last_port"))

        # ID 輸入 (預設 0000)
        ctk.CTkLabel(self.sidebar, text="ID:").pack(pady=(5, 0))
        self.id_entry = ctk.CTkEntry(self.sidebar)
        self.id_entry.pack(pady=5, padx=20)
        self.id_entry.insert(0, "0000")

        # 名稱輸入
        ctk.CTkLabel(self.sidebar, text="名稱:").pack(pady=(5, 0))
        self.prod_combo = ctk.CTkComboBox(self.sidebar, values=self.config.get("product_history", []))
        self.prod_combo.pack(pady=5, padx=20)

        # 瓶數
        ctk.CTkLabel(self.sidebar, text="瓶數:").pack(pady=(5, 0))
        self.bot_entry = ctk.CTkEntry(self.sidebar)
        self.bot_entry.pack(pady=5, padx=20)
        self.bot_entry.insert(0, "1")

        # 項目類型
        self.type_var = ctk.StringVar(value="消耗")
        self.type_seg = ctk.CTkSegmentedButton(self.sidebar, values=["補充", "消耗"], variable=self.type_var)
        self.type_seg.pack(pady=15, padx=20)

        # 開始按鈕
        self.btn_start = ctk.CTkButton(self.sidebar, text="開始監聽", command=self.toggle_monitoring, fg_color="green")
        self.btn_start.pack(pady=20, padx=20)

        # 右側面板
        self.main_content = ctk.CTkFrame(self)
        self.main_content.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.main_content.grid_rowconfigure(2, weight=1)
        self.main_content.grid_columnconfigure(0, weight=1)

        # 大字重顯示 (偵錯模式顯示最後一筆收到的內容)
        self.weight_display = ctk.CTkLabel(self.main_content, text="WAITING", font=ctk.CTkFont(size=60, weight="bold"))
        self.weight_display.pack(pady=20)

        # 關鍵：原始數據日誌區 (加強顯示)
        ctk.CTkLabel(self.main_content, text="--- 原始數據串流 (RAW DATA STREAM) ---", text_color="yellow").pack()
        self.log_box = ctk.CTkTextbox(self.main_content, height=250, font=("Consolas", 12))
        self.log_box.pack(pady=10, padx=20, fill="x")

        # 最近紀錄表格
        self.setup_treeview()

    def setup_treeview(self):
        columns = ("日期", "時間", "ID", "名稱", "類別", "重量", "瓶", "餘量")
        self.tree = ttk.Treeview(self.main_content, columns=columns, show="headings", height=8)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=80, anchor="center")
        self.tree.pack(pady=10, padx=20, fill="both", expand=True)

    # --- 核心偵錯邏輯 ---
    def listen_serial(self):
        while self.is_monitoring:
            try:
                if self.ser.in_waiting > 0:
                    # 讀取原始位元組並嘗試解碼
                    raw_bytes = self.ser.readline()
                    raw_str = raw_bytes.decode('ascii', errors='replace').strip()
                    
                    # 立即在螢幕顯示「所有」收到的內容
                    timestamp = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
                    debug_msg = f"[{timestamp}] RECV >> {repr(raw_bytes)}\n解碼: {raw_str}"
                    
                    self.after(0, lambda m=debug_msg: self.update_log(m))
                    
                    # 嘗試進行原本的資料處理
                    self.after(0, self.handle_weight_data, raw_str)
            except Exception as e:
                self.after(0, lambda: self.update_log(f"連線中斷: {e}"))
                break

    def handle_weight_data(self, raw_line):
        """原有的處理邏輯，但加上失敗原因的日誌"""
        parts = raw_line.split()
        
        # 如果格式不符合原定的 S 或 SD，在 log 提醒
        if not (len(parts) >= 4 and parts[0] in ['S', 'SD']):
            if raw_line:
                self.update_log(f"⚠️ 格式不符過濾條件 (需開頭為 S/SD 且長度>=4)。目前拆解結果: {parts}")
            return

        # 若符合，則執行原本的顯示與存檔
        weight = parts[2]
        unit = parts[3]
        self.weight_display.configure(text=weight)
        
        # 這裡可以串接原本的 process_save ... (為了簡化 debug 僅顯示)
        self.update_log(f"✅ 成功辨識數據: 重量={weight}, 單位={unit}")

    # --- 其餘輔助功能 (維持原狀) ---
    def update_log(self, message):
        self.log_box.insert("end", message + "\n" + "-"*30 + "\n")
        self.log_box.see("end")

    def get_ports(self):
        return [port.device for port in serial.tools.list_ports.comports()] or ["未偵測到裝置"]

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f: return json.load(f)
            except: pass
        return {"last_port": "", "product_history": []}

    def toggle_monitoring(self):
        if not self.is_monitoring:
            port = self.port_menu.get()
            try:
                self.ser = serial.Serial(port, 9600, timeout=1)
                self.is_monitoring = True
                self.btn_start.configure(text="停止監聽", fg_color="red")
                self.update_log(f"開始監聽 {port}...")
                threading.Thread(target=self.listen_serial, daemon=True).start()
            except Exception as e:
                messagebox.showerror("錯誤", f"無法開啟 {port}: {e}")
        else:
            self.is_monitoring = False
            if self.ser: self.ser.close()
            self.btn_start.configure(text="開始監聽", fg_color="green")

if __name__ == "__main__":
    app = BalanceDebugGUI()
    app.mainloop()