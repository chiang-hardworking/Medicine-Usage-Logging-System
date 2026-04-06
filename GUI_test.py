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

class BalanceGUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("METTLER TOLEDO 藥品使用紀錄系統 v3.0")
        self.geometry("950x780") # 加寬加高以容納表格

        # --- 載入設定 ---
        self.config = self.load_config()
        self.file_name = "medicine_data.xlsx"
        self.is_monitoring = False
        self.ser = None

        # --- UI 配置 ---
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ================= 左側控制面板 =================
        self.sidebar = ctk.CTkFrame(self, width=240, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        self.logo_label = ctk.CTkLabel(self.sidebar, text="系統設定", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.pack(pady=15, padx=10)

        # Port 選擇
        self.port_label = ctk.CTkLabel(self.sidebar, text="選擇 COM Port:")
        self.port_label.pack(pady=(5, 0))
        available_ports = self.get_ports()
        self.port_menu = ctk.CTkOptionMenu(self.sidebar, values=available_ports)
        self.port_menu.pack(pady=5, padx=20)
        
        last_port = self.config.get("last_port")
        if last_port in available_ports:
            self.port_menu.set(last_port)

        # ID 輸入 (預設改為 0000)
        self.id_label = ctk.CTkLabel(self.sidebar, text="ID:")
        self.id_label.pack(pady=(5, 0))
        self.id_entry = ctk.CTkEntry(self.sidebar)
        self.id_entry.pack(pady=5, padx=20)
        self.id_entry.insert(0, "0000")

        # 名稱輸入
        self.prod_label = ctk.CTkLabel(self.sidebar, text="名稱:")
        self.prod_label.pack(pady=(5, 0))
        self.prod_combo = ctk.CTkComboBox(self.sidebar, values=self.config.get("product_history", []))
        self.prod_combo.pack(pady=5, padx=20)

        # 瓶數
        self.bot_label = ctk.CTkLabel(self.sidebar, text="瓶數:")
        self.bot_label.pack(pady=(5, 0))
        self.bot_entry = ctk.CTkEntry(self.sidebar)
        self.bot_entry.pack(pady=5, padx=20)
        self.bot_entry.insert(0, "1")

        # 項目類型 (補充 / 消耗)
        self.type_label = ctk.CTkLabel(self.sidebar, text="項目類別:")
        self.type_label.pack(pady=(15, 0))
        self.type_var = ctk.StringVar(value="消耗")
        self.type_seg = ctk.CTkSegmentedButton(self.sidebar, values=["補充", "消耗"], variable=self.type_var)
        self.type_seg.pack(pady=5, padx=20)

        # 修改標準按鈕
        self.btn_mod_std = ctk.CTkButton(self.sidebar, text="修改當前 ID 標準", command=self.modify_standard, fg_color="#b58d35", hover_color="#8a6a25")
        self.btn_mod_std.pack(pady=(20, 0), padx=20)

        # 開始/停止按鈕
        self.btn_start = ctk.CTkButton(self.sidebar, text="開始監聽", command=self.toggle_monitoring, fg_color="green", hover_color="#006400")
        self.btn_start.pack(pady=(30, 20), padx=20)

        # ================= 右側顯示面板 =================
        self.main_content = ctk.CTkFrame(self)
        self.main_content.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.main_content.grid_rowconfigure(2, weight=1) # 讓表格區塊自動延展
        self.main_content.grid_columnconfigure(0, weight=1)

        # 大字重顯示
        self.weight_frame = ctk.CTkFrame(self.main_content, fg_color="transparent")
        self.weight_frame.grid(row=0, column=0, pady=(10, 0))
        self.weight_display = ctk.CTkLabel(self.weight_frame, text="0.0000", font=ctk.CTkFont(size=70, weight="bold"))
        self.weight_display.pack()
        self.unit_label = ctk.CTkLabel(self.weight_frame, text="---", font=ctk.CTkFont(size=25))
        self.unit_label.pack(pady=(0, 10))

        # 日誌
        self.log_box = ctk.CTkTextbox(self.main_content, height=100)
        self.log_box.grid(row=1, column=0, pady=(0, 10), padx=20, sticky="ew")

        # 最近紀錄表格 (Treeview)
        self.setup_treeview()

        # 啟動時載入歷史紀錄
        self.load_recent_records()

    # --- 表格 (Treeview) 設定 ---
    def setup_treeview(self):
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2b2b2b", foreground="white", fieldbackground="#2b2b2b", borderwidth=0, rowheight=25)
        style.configure("Treeview.Heading", background="#3b3b3b", foreground="white", relief="flat", font=('Arial', 10, 'bold'))
        style.map("Treeview", background=[('selected', '#1f538d')])

        self.tree_frame = ctk.CTkFrame(self.main_content)
        self.tree_frame.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="nsew")
        self.tree_frame.grid_columnconfigure(0, weight=1)
        self.tree_frame.grid_rowconfigure(0, weight=1)

        columns = ("日期", "時間", "ID", "名稱", "類別", "重量", "瓶", "餘量")
        self.tree = ttk.Treeview(self.tree_frame, columns=columns, show="headings", height=10)
        
        # 設定欄位寬度
        widths = [90, 80, 60, 100, 50, 80, 40, 80]
        for col, w in zip(columns, widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor="center")
            
        self.tree.grid(row=0, column=0, sticky="nsew")
        
        scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=0, column=1, sticky="ns")

    def load_recent_records(self):
        # 清空現有表格
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        if not os.path.exists(self.file_name):
            return
            
        try:
            wb = load_workbook(self.file_name, read_only=True)
            if "詳細記錄" in wb.sheetnames:
                ws = wb["詳細記錄"]
                rows = list(ws.iter_rows(values_only=True))
                if len(rows) > 1:
                    # 取最後 10 筆，不含標題列
                    recent_data = rows[1:][-10:]
                    # 反轉順序讓最新的一筆在最上面
                    for r in reversed(recent_data):
                        # r 格式: [日期, 時間, ID, 名稱, 類別, 重量, 單位, 瓶數, 標準, 餘量]
                        # 依照 Treeview columns 抓取對應資料 (根據 save_to_excel 的排列)
                        if len(r) >= 10:
                            display_row = (r[0], r[1], r[2], r[3], r[4], f"{r[5]} {r[6]}", r[7], round(float(r[9]), 4) if r[9] else 0)
                            self.tree.insert("", "end", values=display_row)
            wb.close()
        except Exception as e:
            self.update_log(f"載入紀錄失敗: {e}")

    # --- 設定與資料讀取 ---
    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f: return json.load(f)
            except: pass
        return {"last_port": "", "product_history": []}

    def save_config(self):
        self.config["last_port"] = self.port_menu.get()
        current_prod = self.prod_combo.get()
        if current_prod and current_prod not in self.config["product_history"]:
            self.config["product_history"].append(current_prod)
            self.config["product_history"] = self.config["product_history"][-20:]
            self.prod_combo.configure(values=self.config["product_history"])
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.config, f, indent=4, ensure_ascii=False)

    def get_ports(self):
        ports = [port.device for port in serial.tools.list_ports.comports()]
        return ports if ports else ["未偵測到裝置"]

    def get_standard_from_excel(self, cid):
        """從庫存餘量表中查詢指定 ID 的標準"""
        if not os.path.exists(self.file_name): return None
        try:
            wb = load_workbook(self.file_name, read_only=True)
            if "庫存餘量" in wb.sheetnames:
                ws = wb["庫存餘量"]
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if str(row[0]) == str(cid):
                        wb.close()
                        return float(row[3]) if row[3] is not None else 0.0
            wb.close()
        except: pass
        return None

    # --- 邏輯處理 ---
    def toggle_monitoring(self):
        if not self.is_monitoring: self.start_monitoring()
        else: self.stop_monitoring()

    def start_monitoring(self):
        port = self.port_menu.get()
        if port == "未偵測到裝置":
            self.update_log("錯誤: 請先連接天平")
            return
        try:
            self.ser = serial.Serial(port, 9600, timeout=1)
            self.is_monitoring = True
            self.save_config()
            self.btn_start.configure(text="停止監聽", fg_color="red")
            self.update_log(f"系統啟動中... ({port})")
            threading.Thread(target=self.listen_serial, daemon=True).start()
        except Exception as e:
            self.update_log(f"連線失敗: {e}")

    def stop_monitoring(self):
        self.is_monitoring = False
        if self.ser: self.ser.close()
        self.btn_start.configure(text="開始監聽", fg_color="green")
        self.update_log("監聽已停止。")

    def listen_serial(self):
        while self.is_monitoring:
            try:
                if self.ser.in_waiting > 0:
                    raw_line = self.ser.readline().decode('ascii', errors='replace').strip()
                    if raw_line: 
                        # 使用 self.after 切換回主執行緒處理資料，確保 UI 彈出視窗不會卡死
                        self.after(0, self.handle_weight_data, raw_line)
            except: break

    def handle_weight_data(self, raw_line):
        parts = raw_line.split()
        if not (len(parts) >= 4 and parts[0] in ['S', 'SD']):
            if raw_line:
                self.update_log(f"⚠️ 格式不符過濾條件 (需開頭為 S/SD 且長度>=4)。目前拆解結果: {parts}")
            return
    
        weight = parts[2]
        unit = parts[3]
        
        self.weight_display.configure(text=weight)
        self.unit_label.configure(text=unit)
        
        cid = self.id_entry.get()
        prod = self.prod_combo.get()
        ctype = self.type_var.get()
        bot = self.bot_entry.get()

        # 檢查 Excel 中是否有該 ID 的標準
        std_val = self.get_standard_from_excel(cid)
        
        if std_val is None:
            # 第一次出現，彈出輸入視窗
            self.popup_standard_input(cid, prod, ctype, weight, unit, bot, default_std=weight)
        else:
            # 已存在，直接存檔
            self.process_save(cid, prod, ctype, weight, unit, bot, std_val)

    def update_log(self, message):
        self.log_box.insert("end", message + "\n" + "-"*30 + "\n")
        self.log_box.see("end")
    
    def popup_standard_input(self, cid, prod, ctype, weight, unit, bot, default_std):
        """彈出視窗詢問標準"""
        dialog = ctk.CTkToplevel(self)
        dialog.title("新增藥品標準")
        dialog.geometry("350x200")
        dialog.attributes("-topmost", True)
        dialog.grab_set() # 鎖定主視窗
        
        lbl = ctk.CTkLabel(dialog, text=f"ID [{cid}] 首次出現！\n請輸入安全庫存標準:", font=ctk.CTkFont(size=16))
        lbl.pack(pady=(20, 10))
        
        entry = ctk.CTkEntry(dialog, width=150, font=ctk.CTkFont(size=16))
        entry.pack(pady=5)
        entry.insert(0, str(default_std))
        
        def confirm():
            val = entry.get()
            dialog.destroy()
            try: float_val = float(val)
            except: float_val = float(default_std) # 防呆
            self.process_save(cid, prod, ctype, weight, unit, bot, float_val)
            
        btn = ctk.CTkButton(dialog, text="確認寫入", command=confirm)
        btn.pack(pady=15)

    def modify_standard(self):
        """手動修改標準按鈕的邏輯"""
        cid = self.id_entry.get()
        std_val = self.get_standard_from_excel(cid)
        
        if std_val is None:
            messagebox.showwarning("警告", f"ID [{cid}] 尚未建立紀錄！\n請先進行一次秤重來建立資料。")
            return
            
        dialog = ctk.CTkToplevel(self)
        dialog.title("修改藥品標準")
        dialog.geometry("350x200")
        dialog.attributes("-topmost", True)
        dialog.grab_set()
        
        lbl = ctk.CTkLabel(dialog, text=f"目前 ID [{cid}] 的標準為: {std_val}\n請輸入新標準:", font=ctk.CTkFont(size=14))
        lbl.pack(pady=(20, 10))
        
        entry = ctk.CTkEntry(dialog, width=150)
        entry.pack(pady=5)
        entry.insert(0, str(std_val))
        
        def confirm():
            val = entry.get()
            dialog.destroy()
            try: 
                new_std = float(val)
                self.update_standard_only(cid, new_std)
            except:
                messagebox.showerror("錯誤", "請輸入有效的數字！")
            
        btn = ctk.CTkButton(dialog, text="確認修改", command=confirm)
        btn.pack(pady=15)

    def update_standard_only(self, cid, new_std):
        """僅修改標準並更新塗色，不新增秤重紀錄"""
        try:
            wb = load_workbook(self.file_name)
            ws_inv = wb["庫存餘量"]
            alert_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
            normal_fill = PatternFill(fill_type=None)

            for row in range(2, ws_inv.max_row + 1):
                if str(ws_inv.cell(row=row, column=1).value) == str(cid):
                    ws_inv.cell(row=row, column=4, value=new_std)
                    current_rem = float(ws_inv.cell(row=row, column=5).value)
                    
                    # 重新評估顏色
                    current_fill = alert_fill if current_rem < new_std else normal_fill
                    for col in range(1, 6):
                        ws_inv.cell(row=row, column=col).fill = current_fill
                    break
            wb.save(self.file_name)
            self.update_log(f"--- ID [{cid}] 標準已成功修改為 {new_std} ---")
        except PermissionError:
            self.update_log("!!! 修改失敗: 請先關閉 Excel 檔案 !!!")

    def process_save(self, cid, prod, ctype, weight, unit, bot, std_val):
        """執行最終儲存動作"""
        now = datetime.datetime.now()
        date_str = now.strftime("%Y-%m-%d")
        time_str = now.strftime("%H:%M:%S")
        
        log_msg = f"[{time_str}] {ctype} | {prod}({cid}) x{bot}瓶: {weight} {unit} (標準:{std_val})"
        self.update_log(log_msg)
        
        self.save_to_excel(cid, prod, ctype, date_str, time_str, weight, unit, bot, std_val)
        self.save_config()
        self.load_recent_records() # 儲存後立刻更新介面上的表格

    def save_to_excel(self, cid, prod, ctype, d, t, w, u, bot_str, std_val):
        try:
            try: w_val = float(w)
            except: w_val = 0.0
            try: bottles = int(bot_str)
            except: bottles = 1

            alert_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
            normal_fill = PatternFill(fill_type=None)

            if not os.path.exists(self.file_name):
                wb = Workbook()
                ws_log = wb.active
                ws_log.title = "詳細記錄"
                ws_log.append(["日期", "時間", "ID", "名稱", "類別", "重量", "單位", "瓶數", "標準", "餘量"])
                
                ws_inv = wb.create_sheet(title="庫存餘量")
                ws_inv.append(["ID", "名稱", "單位", "標準", "餘量"])
                wb.save(self.file_name)

            wb = load_workbook(self.file_name)
            
            if "詳細記錄" in wb.sheetnames: ws_log = wb["詳細記錄"]
            else:
                ws_log = wb.create_sheet("詳細記錄")
                ws_log.append(["日期", "時間", "ID", "名稱", "類別", "重量", "單位", "瓶數", "標準", "餘量"])

            if "庫存餘量" in wb.sheetnames: ws_inv = wb["庫存餘量"]
            else:
                ws_inv = wb.create_sheet("庫存餘量")
                ws_inv.append(["ID", "名稱", "單位", "標準", "餘量"])

            match_row = None
            current_rem = 0.0
            
            for row in range(2, ws_inv.max_row + 1):
                if str(ws_inv.cell(row=row, column=1).value) == str(cid):
                    match_row = row
                    val = ws_inv.cell(row=row, column=5).value
                    if val is not None:
                        try: current_rem = float(val)
                        except: pass
                    break
            
            total_change = w_val * bottles
            if ctype == "消耗": total_change = -total_change
            new_rem = current_rem + total_change
            
            if match_row:
                ws_inv.cell(row=match_row, column=2, value=prod)
                ws_inv.cell(row=match_row, column=3, value=u)
                ws_inv.cell(row=match_row, column=4, value=std_val)
                ws_inv.cell(row=match_row, column=5, value=new_rem)
            else:
                ws_inv.append([cid, prod, u, std_val, new_rem])
                match_row = ws_inv.max_row
                
            current_fill = alert_fill if new_rem < std_val else normal_fill
            for col in range(1, 6):
                ws_inv.cell(row=match_row, column=col).fill = current_fill

            ws_log.append([d, t, cid, prod, ctype, w_val, u, bottles, std_val, new_rem])
            wb.save(self.file_name)
            
        except PermissionError:
            self.update_log("!!! 存檔失敗: 請先關閉 Excel 檔案 !!!")
        except Exception as e:
            self.update_log(f"存檔發生錯誤: {e}")

if __name__ == "__main__":
    app = BalanceGUI()
    app.mainloop()