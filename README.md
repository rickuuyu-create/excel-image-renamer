# Excel Image Renamer

Batch rename image files from Chinese names to English names using an Excel mapping table.

This project was created for people who need to manage many image assets, teaching materials, recycling/waste-sorting icons, product icons, or dataset images. Instead of renaming files one by one, you can maintain a simple Excel file and let the script rename everything automatically.

> Example: `01_報紙.png` -> `Newspapers.png`

## Why this tool exists

When a folder contains dozens or hundreds of images with local-language names, manual renaming is slow and error-prone. This tool gives you a safer workflow:

1. Put Chinese names in Excel column A.
2. Put English file names in Excel column B.
3. Put images into the `image` folder.
4. Run preview mode first.
5. Check the generated reports.
6. Apply the rename only after confirming the preview.

## Features

- Rename images using an Excel mapping file.
- Preview mode by default, so files are not renamed accidentally.
- Supports blank rows in the Excel file.
- Supports optional header rows.
- Supports common image formats: `.png`, `.jpg`, `.jpeg`, `.webp`, `.bmp`, `.gif`.
- Removes number prefixes such as `01_`, `03-`, or `08 ` from image names before matching.
- Handles Chinese notes in brackets, for example `堅果殼（如核桃殼）` can match `34_堅果殼.png`.
- Generates CSV reports for checking missing, unmatched, duplicate, or incomplete items.
- Windows-friendly batch files included for beginners.

## Folder structure

```text
excel-image-renamer/
├─ rename_images.py
├─ mapping.xlsx
├─ requirements.txt
├─ 1_install_requirements.bat
├─ 2_preview.bat
├─ 3_apply_rename.bat
├─ image/
│  └─ PUT_IMAGES_HERE.txt
└─ rename_reports/          # generated after running the script
```

## Excel format

Use the included `mapping.xlsx`, or create your own file with this layout:

| Column A | Column B |
|---|---|
| 報紙 | Newspapers |
| 剩菜剩飯 | Leftovers |
| 過期藥品 | Expired medicines |
| 陶瓷碎片 | Ceramic shards |

Blank rows are allowed.

## Quick start for Windows beginners

### 1. Install Python

Install Python 3 from the official Python website. During installation, tick the option that says **Add Python to PATH**.

### 2. Install requirements

Double-click:

```text
1_install_requirements.bat
```

### 3. Put images into the image folder

Put your image files into:

```text
image/
```

Example:

```text
image/01_報紙.png
image/03_剩菜剩飯.png
image/08_陶瓷碎片.png
```

### 4. Preview first

Double-click:

```text
2_preview.bat
```

This will not rename your files yet. It only creates reports in:

```text
rename_reports/
```

Check:

```text
rename_reports/renamed_report.csv
```

### 5. Apply rename

After confirming the preview report, double-click:

```text
3_apply_rename.bat
```

Your images will be renamed.

## Command line usage

Preview only:

```bash
python rename_images.py --excel mapping.xlsx --image-dir image
```

Actually rename files:

```bash
python rename_images.py --excel mapping.xlsx --image-dir image --apply
```

Use underscores instead of spaces:

```bash
python rename_images.py --excel mapping.xlsx --image-dir image --apply --use-underscores
```

Use a specific Excel sheet:

```bash
python rename_images.py --excel mapping.xlsx --sheet Sheet1 --image-dir image
```

## Generated reports

The script creates reports inside `rename_reports/`:

| Report | Meaning |
|---|---|
| `renamed_report.csv` | Files that can be renamed or were renamed. |
| `unmatched_images.csv` | Images found in the folder but not found in Excel. |
| `missing_images.csv` | Excel rows that do not have a matching image. |
| `incomplete_excel_rows.csv` | Rows where only Chinese or only English is filled. |
| `duplicate_chinese_rows.csv` | Duplicate Chinese names after normalization. |
| `summary.txt` | Summary of the run. |

## Safety notes

- Run preview mode first.
- Keep a backup of your images before applying rename.
- Do not put private files, passwords, or API keys in this repository.
- If you upload this project publicly, consider replacing sample files with non-sensitive examples.

## Use cases

- Waste sorting education images
- Recycling icon datasets
- Classroom flashcards
- Product image localization
- Dataset cleaning
- Translation workflows
- File naming standardization

## Project status

This is a small beginner-friendly utility. Contributions, suggestions, and bug reports are welcome.

## License

MIT License. See `LICENSE` for details.


# Excel 圖片重新命名工具

使用 Excel 對應表，批次將圖片檔案的中文名稱重新命名為英文名稱。

本專案專為需要管理大量圖片資源（例如教學材料、回收/垃圾分類圖示、產品圖示或資料集圖片）的使用者而設計。無需逐個重命名文件，只需維護一個簡單的 Excel 文件，腳本即可自動完成所有重新命名操作。

> 範例：`01_報紙.png` -> `Newspapers.png`

## 此工具的用途

當一個資料夾包含數十張甚至數百張帶有本地語言名稱的圖片時，手動重命名既緩慢又容易出錯。此工具為您提供更安全的工作流程：

1. 將中文名稱放入 Excel 的 A 欄位。

2. 將英文檔名放入 Excel 的 B 欄。

3. 將圖片放入 `image` 資料夾。

4. 首先運行預覽模式。

5. 檢查產生的報告。

6. 確認預覽後再套用重新命名。

## 功能

- 使用 Excel 對應檔案重新命名影像。

- 預設啟用預覽模式，避免意外重新命名檔案。

- 支援 Excel 檔案中的空白行。

- 支援可選的標題行。

- 支援常見的圖片格式：`.png`、`.jpg`、`.jpeg`、`.webp`、`.bmp`、`.gif`。

- 匹配前移除影像名稱中的數字前綴，例如 `01_`、`03-` 或 `08`。

- 處理括號中的中文註釋，例如 `堅果殼（如核桃殼）` 可以匹配 `34_堅果殼.png`。

- 產生 CSV 報告，用於檢查缺失、未匹配、重複或不完整的項目。

- 包含適用於 Windows 系統的批次文件，方便初學者使用。

## 資料夾結構

```text

excel-image-renamer/

├─ rename_images.py

├─ mapping.xlsx

├─ requirements.txt

├─ 1_install_requirements.bat

├─ 2_preview.bat

├─ 3_apply_rename.bat

├─ image/

│ └─ PUT_IMAGES_HERE.txt

└─ rename_reports/ # 腳本運行後生成

```

## Excel 格式

使用隨附的 `mapping.xlsx` 文件，或建立您自己的文件，並採用以下佈局：

| A 列 | B 列 |

|---|---|

| 報 | Newspapers |

| 剩菜剩飯 | Leftovers |

| 過期藥 | Expired medicines |

| 陶瓷碎片 | Ceramic shards |

允許留空行。

## Windows 新手快速入門

### 1. 安裝 Python

從 Python 官方網站安裝 Python 3。安裝過程中，勾選「將 Python 新增至 PATH」選項。

### 2. 安裝依賴項

按兩下:

```text

1_install_requirements.bat

```

### 3. 將圖片放入 image 資料夾

將圖片檔案放入：

```text

image/

```

範例：

```text

image/01_報紙.png

image/03_剩菜剩話.png

image/08_陶瓷碎片.png

```

### 4. 預覽

按兩下:

```text

2_preview.bat

```

此操作不會立即重新命名您的檔案。它只會在以下位置建立報表：

```text

rename_reports/

```

檢查：

```text

rename_reports/renamed_report.csv

```

### 5. 應用重新命名

確認預覽報告後，按兩下：

```text

3_apply_rename.bat

```

您的圖像將被重新命名。

## 命令列用法

僅供預覽：

『`bash

python rename_images.py --excel mapping.xlsx --image-dir image

```

實際重命名檔案：

『`bash

python rename_images.py --excel mapping.xlsx --image-dir image --apply

```

使用底線代替空格：

『`bash

python rename_images.py --excel mapping.xlsx --image-dir image --apply --use-underscores

```

指定 Excel 工作表：

『`bash

python rename_images.py --excel mapping.xlsx --sheet Sheet1 --image-dir image

```

## 產生的報告

腳本會在 `rename_reports/` 目錄下產生報表：

| 報告 | 意義 |

|---|---|

| `renamed_report.csv` |可重新命名或已重新命名的檔案。 |

| `unmatched_images.csv` | 資料夾中找到但在 Excel 中找不到的圖片。 |

| `missing_images.csv` | Excel 中缺少符合圖片的行。 |

| `incomplete_excel_rows.csv` | 只填寫中文或僅填寫英文的行。 |

| `duplicate_chinese_rows.csv` | 規範化後的重複中文名稱。 |

| `summary.txt` | 運行摘要。 |

## 安全性提示

- 請先執行預覽模式。

- 重新命名前請備份圖片。

- 請勿將私人文件、密碼或 API 金鑰放入此儲存庫。

- 如果您公開上傳此項目，請考慮將範例檔案替換為非敏感範例。

## 應用案例

- 垃圾分類教育圖片

- 回收圖標資料集

- 課堂閃卡

- 產品圖片在地化

- 資料集清洗

- 翻譯工作流程

- 文件命名規範化

## 專案狀態

這是一個面向初學者的小型實用工具。歡迎貢獻代碼、提出建議和提交錯誤報告。

## 許可證

MIT 許可證。詳情請參閱 `LICENSE` 文件。

