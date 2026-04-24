# 班表圖片轉 Excel

這個工具會把月班表圖片中的指定姓名擷取出來，輸出成個人 Excel 班表。

## 功能

- 讀取班表圖片中的年、月、日期
- 擷取指定姓名的早班、午班、晚班
- 自動填入時數：
  - 早班 `4`
  - 午班 `5`
  - 晚班 `4`
- 產出欄位：`年`、`月`、`日`、`周幾`、`班別`、`時數`、`時薪`、`單量`、`總薪水`

## 支援平台

- Windows：可執行 `.exe`
- macOS：可直接用 Python 執行，或在 Mac 上打包成 `.app`

注意：`Windows 的 EXE 不能直接在 Mac 上執行`。

## 直接執行

Windows 或 macOS 都可以：

```bash
python schedule_to_excel.py
```

也可以用命令列模式：

```bash
python schedule_to_excel.py --input "/path/to/schedule.jpg" --name "張盈慧" --output "/path/to/result.xlsx"
```

## 安裝套件

```bash
python -m pip install rapidocr_onnxruntime opencv-python-headless openpyxl pyinstaller
```

## Windows 打包

```powershell
build_exe.bat
```

打包完成後，執行檔會在 `dist\schedule_to_excel.exe`。

## macOS 打包

請在 Mac 電腦上執行：

```bash
chmod +x build_macos.sh
./build_macos.sh
```

打包完成後，程式會在：

```bash
dist/schedule_to_excel.app
```

如果只想直接執行，不打包也可以：

```bash
python schedule_to_excel.py
```
