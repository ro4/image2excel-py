import tkinter as tk
from tkinter import filedialog, messagebox
from image2excel import image_to_excel

class ImageConverterApp:
    def __init__(self, master):
        self.master = master
        master.title("Image to Excel Converter")
        master.geometry("400x200")

        # 创建界面组件
        self.input_btn = tk.Button(master, text="Select Image", command=self.select_image)
        self.output_btn = tk.Button(master, text="Save Path", command=self.select_output)
        self.convert_btn = tk.Button(master, text="Start Conversion", command=self.convert,
                                   state=tk.DISABLED)

        # 路径显示标签
        self.input_label = tk.Label(master, text="No image selected")
        self.output_label = tk.Label(master, text="No save path specified")

        # 布局组件
        self.input_btn.pack(pady=5)
        self.input_label.pack()
        self.output_btn.pack(pady=5)
        self.output_label.pack()
        self.convert_btn.pack(pady=20)

        # 初始化路径变量
        self.image_path = ""
        self.output_path = ""

    def select_image(self):
        filetypes = [
            ('Image files', '.jpg .jpeg .png .bmp'),
            ('所有文件', '*.*')
        ]
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            self.image_path = path
            self.input_label.config(text=path)
            self.check_ready()

    def select_output(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[('Excel File', '.xlsx')]
        )
        if path:
            self.output_path = path
            self.output_label.config(text=path)
            self.check_ready()

    def check_ready(self):
        if self.image_path and self.output_path:
            self.convert_btn.config(state=tk.NORMAL)

    def convert(self):
        try:
            image_to_excel(self.image_path, self.output_path)
            messagebox.showinfo("Success", "File converted successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ImageConverterApp(root)
    root.mainloop()