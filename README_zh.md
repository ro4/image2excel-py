# image2excel

Python工具，将图像转换为Excel电子表格，通过单元格着色

## 功能
- 将图像像素转换为Excel单元格
- 保持颜色准确性
- 调整单元格尺寸以创建方形单元格
- 支持多种图像格式（JPEG、PNG等）

## 要求
- Python 3.6+
- Pillow
- openpyxl

## 安装
```bash
pip install -r requirements.txt
```

## 使用方法

```bash
python image2excel.py <输入图片> <输出Excel文件>
```

### 示例

```bash
python image2excel.py input.png output.xlsx
```

### 图形界面使用方法

1. 启动图形界面：
```bash
python gui.py
```
2. 通过文件浏览器选择输入图片
3. 选择输出的Excel文件路径
4. 点击"转换"按钮生成表格文件



### 参数
- `输入图片`: 输入图片文件路径
- `输出Excel文件`: 输出Excel文件路径

## 注意事项
- 处理大图可能需要较长时间
- 生成的Excel文件可能比原图大很多
- 建议使用较小的图像以获得更好性能