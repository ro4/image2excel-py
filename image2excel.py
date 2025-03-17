from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import sys

def rgb_to_hex(rgb):
    return '%02x%02x%02x' % rgb

def image_to_excel(image_path, output_path):
    # Open image file
    img = Image.open(image_path)
    
    # Get image dimensions
    width, height = img.size
    
    # Create new Excel workbook
    wb = Workbook()
    ws = wb.active
    
    # Set column width and row height to make cells square-like (1 char width ≈8px, 8pt height ≈12px)
    for x in range(width):
        ws.column_dimensions[ws.cell(row=1, column=x+1).column_letter].width = 1
    
    for y in range(height):
        ws.row_dimensions[y+1].height = 6  # 6点≈8像素（1点=1.333像素）
    
    # Iterate through image pixels
    for y in range(height):
        # 设置行高（6点≈8像素）
        ws.row_dimensions[y+1].height = 6
        
        for x in range(width):
            # Get pixel RGB value
            pixel = img.getpixel((x, y))
            if isinstance(pixel, int):
                # Handle grayscale image
                pixel = (pixel, pixel, pixel)
            elif len(pixel) == 4:
                # Handle RGBA image, remove alpha channel
                pixel = pixel[:3]
            
            # Convert RGB to hex color code
            hex_color = rgb_to_hex(pixel)
            
            # Set cell background color
            cell = ws.cell(row=y+1, column=x + 1)
            cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type='solid')
    
    # 保存Excel文件
    wb.save(output_path)

def main():
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
        print(f"Error: {str(e)}")
        print(f"错误：{str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main()