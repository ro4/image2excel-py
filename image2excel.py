"""
将图片转换为Excel表格，每个单元格对应图片像素颜色
Convert image to Excel spreadsheet with cell colors matching image pixels

功能 Features:
1. 支持常见图片格式（JPG/PNG等）
   Supports common image formats (JPG/PNG/etc.)
2. 自动处理灰度/RGBA格式
   Auto-handles grayscale/RGBA formats
3. 生成等比例的Excel表格
   Generates proportionally scaled Excel sheet
"""

from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import sys
from typing import Tuple

def rgb_to_hex(rgb: Tuple[int, int, int]) -> str:
    """
    将RGB元组转换为十六进制颜色代码
    Convert RGB tuple to hexadecimal color code
    
    Args:
        rgb: 包含3个0-255整数的元组
              Tuple of 3 integers (0-255)
    
    Returns:
        str: 6位十六进制颜色代码
             6-digit hex color code
    """
    return '%02x%02x%02x' % rgb

def image_to_excel(image_path: str, output_path: str) -> None:
    """
    将图片转换为带颜色填充的Excel文件
    Convert image to color-filled Excel file
    
    Args:
        image_path: 输入图片路径
                   Input image path
        output_path: 输出Excel路径
                    Output Excel path
    """
    # 打开图片文件 / Open image file
    img = Image.open(image_path)
    
    # 获取图片尺寸 / Get image dimensions
    width, height = img.size
    
    # 创建Excel工作簿 / Create new Excel workbook
    wb = Workbook()
    ws = wb.active
    
    # 设置列宽行高为正方形单元格 / Set column width & row height for square cells
    # 字符宽度≈8px，6点行高≈8px (1点=1.333px) / 1 char width≈8px, 6pt height≈8px (1pt=1.333px)
    for x in range(width):
        ws.column_dimensions[ws.cell(row=1, column=x+1).column_letter].width = 1
    
    for y in range(height):
        ws.row_dimensions[y+1].height = 6  # 6点≈8像素（1点=1.333像素）
    
    # 遍历图片像素 / Iterate through image pixels
    for y in range(height):
        # 设置行高（6点≈8像素） / Set row height (6 points≈8px)
        ws.row_dimensions[y+1].height = 6
        
        for x in range(width):
            # Get pixel RGB value
            pixel = img.getpixel((x, y))
            if isinstance(pixel, int):
                # 处理灰度图像 / Handle grayscale image
                pixel = (pixel, pixel, pixel)
            elif len(pixel) == 4:
                # 处理RGBA图像，移除透明通道 / Handle RGBA image, remove alpha channel
                pixel = pixel[:3]
            
            # Convert RGB to hex color code
            hex_color = rgb_to_hex(pixel)
            
            # Set cell background color
            cell = ws.cell(row=y+1, column=x + 1)
            cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type='solid')
    
    # 保存Excel文件 / Save Excel file
    wb.save(output_path)

def main() -> None:
    """
    主函数处理命令行参数
    Main function handling command line arguments
    """
    if len(sys.argv) != 3:
        print("Usage: python image2excel.py <input_image> <output_excel>")
        print("用法：python image2excel.py <输入图片> <输出Excel文件>")
        sys.exit(1)
    
    image_path = sys.argv[1]
    output_path = sys.argv[2]
    
    try:
        image_to_excel(image_path, output_path)
        print(f"Successfully converted {image_path} to {output_path}")
        print(f"成功将 {image_path} 转换为 {output_path}")
    except Exception as e:
        # 错误处理 / Error handling
        print(f"Error: {str(e)}")
        print(f"错误：{str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main()