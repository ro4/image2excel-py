# image2excel

Python tool to convert images to Excel spreadsheets with cell coloring

## Features
- Converts image pixels to Excel cells
- Preserves color accuracy
- Adjusts cell dimensions to create square cells
- Supports various image formats (JPEG, PNG, etc.)

[中文文档](README_zh.md)

## Requirements
- Python 3.6+
- Pillow
- openpyxl

## Installation
```bash
pip install -r requirements.txt
```

## Usage / 使用方法

```bash
python image2excel.py <input_image> <output_excel>
```

### Example / 示例：

```bash
python image2excel.py input.png output.xlsx
```

### Parameters / 参数：

- `input_image`: Path to the input image file / 输入图片文件的路径
- `output_excel`: Path for the output Excel file / 输出Excel文件的路径

## Notes / 注意事项

- 大尺寸图片可能需要较长处理时间
- 生成的Excel文件可能比原始图片大很多
- 建议使用较小尺寸的图片以获得更好的性能

- Large images may take longer to process
- Generated Excel files might be much larger than the original image
- Smaller images are recommended for better performance