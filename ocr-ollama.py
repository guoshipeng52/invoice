import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from paddleocr import PaddleOCR
import fitz  # PyMuPDF
import os
import GPUtil
import subprocess
import re
import requests
import json
import pandas as pd
from openpyxl import Workbook, load_workbook
from PIL import Image
import numpy as np
from openpyxl import styles



class InvoiceProcessor:
    def __init__(self):
        # 创建加载窗口
        self.load_window = tk.Tk()
        self.load_window.title("加载模型")
        self.load_window.geometry("300x150")
        
        # 居中显示
        self.load_window.eval('tk::PlaceWindow . center')
        
        # 添加标签
        label = tk.Label(self.load_window, text="请先加载模型", font=('Arial', 12))
        label.pack(pady=20)
        
        # 添加加载按钮
        self.load_button = ttk.Button(
            self.load_window, 
            text="加载 Ollama qwen2.5", 
            command=self.load_model
        )
        self.load_button.pack(pady=10)
        
        # 添加进度标签
        self.load_status = tk.Label(self.load_window, text="")
        self.load_status.pack(pady=10)
        
        # 等待加载窗口关闭后再创建主窗口
        self.load_window.mainloop()

    def load_model(self):
        self.load_button.config(state=tk.DISABLED)
        self.load_status.config(text="正在加载模型...")
        
        try:
            # 测试 Ollama 是否运行并加载模型
            response = requests.post(
                "http://localhost:11434/api/generate",
                json={
                    "model": "qwen2.5:14b",
                    "prompt": "test",
                    "stream": False
                }
            )
            
            if response.status_code == 200:
                self.load_status.config(text="模型加载成功！")
                self.load_window.after(1000, self.start_main_window)
            else:
                raise Exception("模型加载失败")
                
        except Exception as e:
            messagebox.showerror("错误", f"模型加载失败: {str(e)}\n请确保Ollama服务正在运行且已安装qwen2.5模型")
            self.load_button.config(state=tk.NORMAL)
            self.load_status.config(text="加载失败，请重试")

    def start_main_window(self):
        # 关闭加载窗口
        self.load_window.destroy()
        
        # 初始化OCR
        self.ocr = PaddleOCR(use_angle_cls=True, lang='ch')
        
        # 创建主窗口
        self.setup_ui()
        self.root.mainloop()

    def check_gpu_vram(self, required_vram_gb=4):
        gpu_vram_gb = 0
        checked_method = "N/A"
        try:
            gpus = GPUtil.getGPUs()
            if gpus:
                gpu_vram_gb = gpus[0].memoryTotal / 1024  # memoryTotal is in MB
                checked_method = "GPUtil"
            else: # GPUtil ran but found no GPUs, try nvidia-smi
                raise Exception("No GPUs found by GPUtil") 
        except Exception as e_gputil:
            # print(f"GPUtil failed: {e_gputil}, trying nvidia-smi...") # For debugging
            try:
                # For Windows, prevent console window
                cflags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                result = subprocess.run(
                    ['nvidia-smi', '--query-gpu=memory.total', '--format=csv,noheader,nounits'],
                    capture_output=True, text=True, check=True, creationflags=cflags
                )
                # Output is like "8192" (MiB), first GPU if multiple.
                # To handle multiple GPUs, one might need to parse more carefully or average/sum.
                # For simplicity, taking the first line if multiple are returned.
                first_line = result.stdout.strip().split('\n')[0]
                if first_line:
                    gpu_vram_gb = int(first_line) / 1024 # Convert MiB to GB
                    checked_method = "nvidia-smi"
            except Exception as e_nvidia_smi:
                # print(f"nvidia-smi failed: {e_nvidia_smi}") # For debugging
                pass # Both methods failed or GPU not detected

        if gpu_vram_gb > 0 and gpu_vram_gb < required_vram_gb:
            messagebox.showwarning(
                "Hardware Warning",
                f"Detected {gpu_vram_gb:.1f}GB VRAM using {checked_method}. "
                f"This is below the recommended {required_vram_gb}GB. "
                "Performance might be affected or OCR/AI models may fail."
            )
        elif checked_method == "N/A": # Both methods failed to find a value
             messagebox.showinfo(
                "Hardware Info",
                "Could not automatically determine GPU VRAM. "
                "Please ensure your hardware meets model requirements if you encounter issues."
            )
        elif gpu_vram_gb >= required_vram_gb :
            print(f"VRAM check: {gpu_vram_gb:.1f}GB detected via {checked_method}. Meets {required_vram_gb}GB requirement.")
        # If gpu_vram_gb is 0 but one of the methods ran, it means 0 VRAM was reported (e.g. integrated graphics by nvidia-smi)
        # This case is covered by the first conditional if required_vram_gb > 0

    def setup_ui(self):
        self.root = tk.Tk()
        self.root.title("发票信息提取")
        self.root.geometry("400x200")

        # 创建选择文件按钮
        self.select_button = tk.Button(self.root, text="选择PDF文件", command=self.select_pdf)
        self.select_button.pack(pady=20)

        # 创建运行按钮
        self.run_button = tk.Button(self.root, text="运行处理", command=self.process_pdf, state=tk.DISABLED)
        self.run_button.pack(pady=10)

        # 创建状态标签
        self.status_label = tk.Label(self.root, text="等待选择文件...")
        self.status_label.pack(pady=10)

        # 添加文件路径显示
        self.file_label = tk.Label(self.root, text="未选择文件", wraplength=350)
        self.file_label.pack(pady=10)

        # 添加返回值显示标签
        self.result_label = tk.Label(self.root, text="Ollama 返回结果：")
        self.result_label.pack(pady=(10,0))

        # 修改返回值显示文本框
        self.result_text = tk.Text(self.root, height=15, width=60)
        self.result_text.pack(pady=5, padx=10)
        
        # 添加滚动条
        scrollbar = tk.Scrollbar(self.root, command=self.result_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.config(yscrollcommand=scrollbar.set)

    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update()

    def pdf_to_images(self, pdf_path):
        doc = fitz.open(pdf_path)
        images = []
        for page in doc:
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            images.append(np.array(img))
        return images

    def select_pdf(self):
        # 选择PDF文件
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.pdf_path:
            self.file_label.config(text=f"已选择: {self.pdf_path}")
            self.run_button.config(state=tk.NORMAL)
            self.update_status("文件已选择，请点击运行处理")
        else:
            self.file_label.config(text="未选择文件")
            self.run_button.config(state=tk.DISABLED)
            self.update_status("等待选择文件...")

    def is_scanned_pdf(self, pdf_path):
        """检查PDF是否为扫描件"""
        doc = fitz.open(pdf_path)
        text_content = ""
        for page in doc:
            text_content += page.get_text()
        doc.close()
        return len(text_content.strip()) < 100

    def process_pdf(self):
        if not hasattr(self, 'pdf_path') or not self.pdf_path:
            self.update_status("请先选择PDF文件")
            return

        self.update_status("正在处理PDF...")
        self.run_button.config(state=tk.DISABLED)
        
        # 创建工作目录
        work_dir = "working"
        if not os.path.exists(work_dir):
            os.makedirs(work_dir)

        text_content = []
        
        # 判断是否为扫描PDF
        if self.is_scanned_pdf(self.pdf_path):
            self.update_status("检测到扫描PDF，正在进行OCR识别...")
            # 转换PDF到图片并OCR
            images = self.pdf_to_images(self.pdf_path)
            for img in images:
                result = self.ocr.ocr(img)
                for line in result[0]:
                    text_content.append(line[1][0])
        else:
            self.update_status("检测到可复制PDF，正在直接提取文本...")
            doc = fitz.open(self.pdf_path)
            extracted_pages_text = []
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                extracted_pages_text.append(page.get_text("text"))
            doc.close()
            # The current OCR processing results in a list of lines.
            # To maintain consistency for the prompt construction,
            # split the combined text from all pages into lines.
            full_text = "\n".join(extracted_pages_text)
            text_content.extend(full_text.splitlines()) # Use extend to add all lines to the existing list

        # 保存结果到txt
        txt_path = os.path.join(work_dir, "ocr_result.txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write("\n".join(text_content))

        self.update_status("正在调用Ollama API...")

        # 调用Ollama API
        prompt = """
        请从以下文本中按顺序提取这些信息，只返回信息内容，每行一个，不要其他任何文字，其中购买方再左上角，销售方：
        第1行：购买方名称
        第2行：购买方的纳税人识别号
        第3行：购买方开户行
        第4行：购买方开户行账号
        第5行：发票代码
        第6行：货物或应税劳务、服务名称
        第7行：数量（若无，则返回“1”）
        第8行：单价（若无，则返回金额）
        第9行：销售方名称
        第10行：销售方的纳税人识别号
        第11行：销售方开户行
        第12行：销售方开户行账号
        第13行：开票日期（返回格式为“YYYYMMDD”）
        第14行：发票号码（若有No.xxx就返回号码，若无则返回“ ”）
        第15行：价税合计（返回数字格式）
        第16行：税率
        第17行：税额（若无则返回“0”）
        第18行：金额
        第19行：发票类型（只返回“普通发票”或“专用发票”）

        
        文本内容：
        """ + "\n".join(text_content)

        try:
            response = requests.post(
                "http://localhost:11434/api/generate",
                json={
                    "model": "qwen2.5:14b",
                    "prompt": prompt,
                    "stream": False
                }
            )
            
            if response.status_code == 200:
                result = response.json()
                extracted_info = result['response'].strip().split('\n')
                
                # 在文本框中显示格式化的返回值
                self.result_text.delete('1.0', tk.END)
                self.result_text.insert('1.0', "提取结果：\n" + "="*50 + "\n")
                
                titles = [
                    "购买方名称", "购买方纳税人识别号", "购买方开户行", 
                    "购买方开户行账号", "发票编号", "货物或应税劳务、服务名称",
                    "数量", "单价", "销售方名称", "销售方的纳税人识别号",
                    "销售方开户行", "销售方开户行账号"
                ]
                
                for i, (title, info) in enumerate(zip(titles, extracted_info)):
                    self.result_text.insert(tk.END, f"{title}: {info}\n")
                
                self.result_text.insert(tk.END, "="*50)
                
                # 使文本框滚动到顶部
                self.result_text.see("1.0")
                
                # 先创建临时Excel文件并保存ollama返回的数据
                temp_wb = Workbook()
                temp_ws = temp_wb.active
                
                # 将19个返回值写入临时文件第1行
                cells = [f"{chr(65+i)}1" for i in range(19)]  # 生成 A1, B1, C1...R1
                for cell, info in zip(cells, extracted_info):
                    temp_ws[cell] = info.strip()
                
                # 保存临时文件
                temp_file = "temp_data.xlsx"
                temp_wb.save(temp_file)
                
                # 打开模板Excel
                template_path = filedialog.askopenfilename(
                    filetypes=[("Excel files", "*.xlsx")],
                    title="选择Excel模板文件"
                )
                
                if template_path:
                    wb = load_workbook(template_path)
                    ws = wb.active
                    
                    # 只保留前两行，删除其他行
                    for row in range(3, ws.max_row + 1):
                        ws.delete_rows(3)
                    
                    # 取消第三行所有单元格的锁定状态
                    for cell in ws[3]:
                        cell.protection = styles.Protection(locked=False)
                    
                    # 按要求填写数据
                    ws['E3'] = temp_ws['A1'].value  # A1 -> E3
                    ws['F3'] = temp_ws['B1'].value  # B1 -> F3
                    ws['B3'] = temp_ws['E1'].value  # E1 -> B3
                    ws['I3'] = f"{temp_ws['F1'].value}*00010-"  # F1 -> I3 + "*00010-"
                    ws['L3'] = temp_ws['G1'].value  # G1 -> L3
                    ws['M3'] = temp_ws['H1'].value  # H1 -> M3
                    ws['G3'] = temp_ws['I1'].value  # I1 -> G3
                    ws['H3'] = temp_ws['J1'].value  # J1 -> H3
                    ws['D3'] = temp_ws['M1'].value  # M1 -> D3
                    ws['C3'] = temp_ws['N1'].value  # N1 -> C3
                    ws['Q3'] = temp_ws['O1'].value  # O1 -> Q3
                    ws['O3'] = temp_ws['P1'].value  # P1 -> O3
                    ws['P3'] = temp_ws['Q1'].value  # P3 -> Q1
                    ws['N3'] = temp_ws['R1'].value  # R1 -> N3
                    ws['A3'] = temp_ws['S1'].value  # S1 -> A3

                    
                    # 让用户选择保存新Excel文件的位置
                    save_path = filedialog.asksaveasfilename(
                        defaultextension=".xlsx",
                        filetypes=[("Excel files", "*.xlsx")],
                        title="选择保存Excel文件的位置"
                    )
                    
                    if save_path:
                        wb.save(save_path)
                        self.update_status(f"处理完成！结果已保存到: {save_path}")
                        
                        # 创建确认对话框
                        open_dialog = tk.Toplevel(self.root)
                        open_dialog.title("文件已保存")
                        open_dialog.geometry("300x150")
                        
                        # 设置模态
                        open_dialog.transient(self.root)
                        open_dialog.grab_set()
                        
                        # 添加提示文本
                        tk.Label(
                            open_dialog, 
                            text="文件保存成功！\n是否立即打开文件？",
                            font=('Arial', 10)
                        ).pack(pady=20)
                        
                        # 按钮框架
                        btn_frame = tk.Frame(open_dialog)
                        btn_frame.pack(pady=10)
                        
                        def open_file():
                            os.startfile(save_path)
                            open_dialog.destroy()
                            
                        def close_dialog():
                            open_dialog.destroy()
                        
                        # 使用 tk.Button 替代 ttk.Button，并设置固定宽度和样式
                        tk.Button(
                            btn_frame, 
                            text="打开", 
                            command=open_file,
                            width=10,
                            relief="raised",
                            bg="#e1e1e1"
                        ).pack(side=tk.LEFT, padx=10)
                        
                        tk.Button(
                            btn_frame, 
                            text="关闭", 
                            command=close_dialog,
                            width=10,
                            relief="raised",
                            bg="#e1e1e1"
                        ).pack(side=tk.LEFT, padx=10)
                        
                        # 设置窗口居中
                        window_width = 300
                        window_height = 150
                        screen_width = self.root.winfo_screenwidth()
                        screen_height = self.root.winfo_screenheight()
                        x = (screen_width - window_width) // 2
                        y = (screen_height - window_height) // 2
                        open_dialog.geometry(f"{window_width}x{window_height}+{x}+{y}")
                        
                        # 等待对话框关闭
                        self.root.wait_window(open_dialog)
                        
                    else:
                        self.update_status("已取消保存Excel文件")
                    
                    # 删除临时文件
                    os.remove(temp_file)
            else:
                self.update_status("API调用失败")

        except Exception as e:
            self.update_status(f"发生错误: {str(e)}")

        # 处理成后重新启用按钮
        self.run_button.config(state=tk.NORMAL)

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    processor = InvoiceProcessor()
    processor.run()
