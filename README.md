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
