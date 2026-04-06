# Medicine Usage Logging System

實驗室秤重與庫存管理工具，透過 COM Port 讀取天平資料，自動寫入 Excel。

## 主要功能
- 自動監聽天平資料（S/SD 格式）
- Excel 雙工作表管理（`詳細記錄`、`庫存餘量`）
- 首次產品編號自動詢問安全標準
- 低於標準自動標紅
- 內建測試模式（預設隱藏，`Ctrl+Shift+T` 解鎖）
- 內建「檢查更新」按鈕（比對 GitHub Releases 最新版本）

## 開發環境安裝
```powershell
pip install -r requirements.txt
```

## 本機執行
```powershell
python GUI_test.py
```

## 版本檔
- `version.json` 內維護目前版本，例如：
```json
{
  "version": "3.2.0"
}
```

## 打包 EXE（推薦流程）
在專案根目錄執行：
```powershell
.\build_exe.ps1
```

若要在打包時直接覆蓋版本號：
```powershell
.\build_exe.ps1 -Version 3.2.1
```

打包結果：
- 執行檔資料夾：`output/release/`
- 發佈壓縮檔：`output/MedicineUsageLoggingSystem-v<version>.zip`

## 發佈新版（GitHub Releases）
1. 更新程式碼與 `version.json`
2. 執行 `.\build_exe.ps1`
3. 將 `output/MedicineUsageLoggingSystem-v<version>.zip` 上傳到 GitHub Release
4. Release tag 建議使用 `v<version>`（例如 `v3.2.1`）

## 使用者更新方式
1. 開啟程式
2. 按「檢查更新」
3. 若有新版本，按提示前往 GitHub Release 下載
