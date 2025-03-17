# image2excel

Python tool to convert images to Excel spreadsheets with cell coloring

## Features
- Converts image pixels to Excel cells
- Preserves color accuracy
- Adjusts cell dimensions to create square cells
- Supports various image formats (JPEG, PNG, etc.)

**[中文文档](README_zh.md)**

## Requirements
- Python 3.6+
- Pillow
- openpyxl

## Installation
```bash
pip install -r requirements.txt
```

## Usage

```bash
python image2excel.py <input_image> <output_excel>
```

### Example

```bash
python image2excel.py input.png output.xlsx
```

### Parameters

- `input_image`: Path to the input image file
- `output_excel`: Path for the output Excel file

### GUI Usage

1. Start the graphical interface:
```bash
python gui.py
```
2. Select input image file using the file browser
3. Choose output Excel file path
4. Click "Convert" button to generate spreadsheet


## Notes

- Large images may take longer to process
- Generated Excel files might be much larger than the original image
- Smaller images are recommended for better performance