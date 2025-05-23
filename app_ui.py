import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os # For os.startfile in prompt_and_open_file

class AppUI:
    def __init__(self, root, controller):
        self.root = root
        self.controller = controller
        # UI elements will be initialized in setup_main_ui or show_loading_window

    def show_loading_window(self, load_model_command):
        """
        Creates and displays the initial model loading window.
        Args:
            load_model_command: The command to be executed when the load button is pressed.
        """
        self.load_window = tk.Toplevel(self.root) # Use Toplevel if root is main window
        self.load_window.title("加载模型")
        self.load_window.geometry("300x150")
        self.load_window.eval('tk::PlaceWindow . center') # Center it

        # Make loading window modal if root is already defined and withdrawn
        if self.root.winfo_exists() and self.root.state() == 'withdrawn':
             self.load_window.transient(self.root)
             self.load_window.grab_set()


        label = tk.Label(self.load_window, text="请先加载模型", font=('Arial', 12))
        label.pack(pady=20)

        self.load_button = ttk.Button(
            self.load_window,
            text="加载 Ollama qwen2.5",
            command=load_model_command
        )
        self.load_button.pack(pady=10)

        self.load_status_label = tk.Label(self.load_window, text="")
        self.load_status_label.pack(pady=10)
        
        # If the root window is the main app window, we might not start its mainloop here.
        # The main_app.py will handle the main root.mainloop().
        # If self.root is just a dummy Tk() for the loading window, then load_window.mainloop() is fine.
        # For now, assume main_app.py starts the overall mainloop.
        # self.load_window.mainloop() # This would block if called on a Toplevel if root's mainloop isn't running

    def close_loading_window(self):
        if hasattr(self, 'load_window') and self.load_window.winfo_exists():
            self.load_window.destroy()

    def setup_main_ui(self):
        """Sets up the main User Interface for the application."""
        self.root.title("发票信息提取")
        self.root.geometry("400x200") # Adjust as needed, or make dynamic

        self.select_button = tk.Button(self.root, text="选择PDF文件", command=self.controller.select_pdf_action)
        self.select_button.pack(pady=20)

        self.run_button = tk.Button(self.root, text="运行处理", command=self.controller.process_pdf_action, state=tk.DISABLED)
        self.run_button.pack(pady=10)

        self.status_label = tk.Label(self.root, text="等待选择文件...")
        self.status_label.pack(pady=10)

        self.file_label = tk.Label(self.root, text="未选择文件", wraplength=350)
        self.file_label.pack(pady=10)

        self.result_label = tk.Label(self.root, text="Ollama 返回结果：")
        self.result_label.pack(pady=(10,0))

        self.result_text_area = tk.Text(self.root, height=15, width=60)
        self.result_text_area.pack(pady=5, padx=10)
        
        scrollbar = tk.Scrollbar(self.root, command=self.result_text_area.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text_area.config(yscrollcommand=scrollbar.set)
        
        if self.root.state() == 'withdrawn': # If root was hidden for loading screen
            self.root.deiconify() # Show the main window

    def update_status(self, message):
        self.status_label.config(text=message)
        # self.root.update() # Controller should decide when to force update if needed

    def update_file_label(self, filepath):
        self.file_label.config(text=f"已选择: {filepath}")

    def enable_run_button(self, enabled=True):
        self.run_button.config(state=tk.NORMAL if enabled else tk.DISABLED)

    def show_results(self, content_json_str):
        self.result_text_area.delete('1.0', tk.END)
        self.result_text_area.insert('1.0', "提取结果 (JSON):\n" + "="*50 + "\n")
        self.result_text_area.insert(tk.END, content_json_str) # Assumes pretty-printed JSON
        self.result_text_area.insert(tk.END, "\n" + "="*50)
        self.result_text_area.see("1.0")
        
    def show_error_in_results(self, error_message):
        self.result_text_area.delete('1.0', tk.END)
        self.result_text_area.insert('1.0', error_message)

    def prompt_and_open_file(self, file_path):
        """Creates a dialog asking to open the saved file."""
        open_dialog = tk.Toplevel(self.root)
        open_dialog.title("文件已保存")
        open_dialog.geometry("300x150")
        
        open_dialog.transient(self.root)
        open_dialog.grab_set()
        
        tk.Label(
            open_dialog, 
            text="文件保存成功！\n是否立即打开文件？",
            font=('Arial', 10)
        ).pack(pady=20)
        
        btn_frame = tk.Frame(open_dialog)
        btn_frame.pack(pady=10)
        
        def do_open_file():
            try:
                os.startfile(file_path)
            except AttributeError: # os.startfile is not on all OSes (e.g. some Linux minimal installs)
                messagebox.showinfo("打开文件", f"请手动打开文件: {file_path}", parent=open_dialog)
            except Exception as e:
                messagebox.showerror("打开文件失败", f"无法打开文件: {e}", parent=open_dialog)
            open_dialog.destroy()
            
        tk.Button(
            btn_frame, text="打开", command=do_open_file,
            width=10, relief="raised", bg="#e1e1e1"
        ).pack(side=tk.LEFT, padx=10)
        
        tk.Button(
            btn_frame, text="关闭", command=open_dialog.destroy,
            width=10, relief="raised", bg="#e1e1e1"
        ).pack(side=tk.LEFT, padx=10)
        
        # Center dialog
        self.root.eval(f'tk::PlaceWindow {str(open_dialog)} center')
        self.root.wait_window(open_dialog)

    def update_loading_status(self, message):
        if hasattr(self, 'load_status_label') and self.load_status_label.winfo_exists():
            self.load_status_label.config(text=message)
            # self.load_window.update() # Might not be necessary, controller can manage updates

    def enable_load_button(self, enabled=True):
        if hasattr(self, 'load_button') and self.load_button.winfo_exists():
            self.load_button.config(state=tk.NORMAL if enabled else tk.DISABLED)

    # Wrappers for tkinter dialogs, so controller doesn't directly use filedialog/messagebox
    def ask_open_pdf_file(self):
        return filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])

    def ask_excel_template_file(self):
        return filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            title="选择Excel模板文件"
        )

    def ask_save_excel_file(self):
        return filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="选择保存Excel文件的位置"
        )

    def show_error_message(self, title, message):
        messagebox.showerror(title, message, parent=self.root if self.root.winfo_exists() and self.root.state() != 'withdrawn' else None)

    def show_info_message(self, title, message):
        messagebox.showinfo(title, message, parent=self.root if self.root.winfo_exists() and self.root.state() != 'withdrawn' else None)

    def show_warning_message(self, title, message):
        messagebox.showwarning(title, message, parent=self.root if self.root.winfo_exists() and self.root.state() != 'withdrawn' else None)

# Example of how it might be used (for testing this file standalone)
if __name__ == '__main__':
    class MockController:
        def select_pdf_action(self): print("Controller: select_pdf_action called")
        def process_pdf_action(self): print("Controller: process_pdf_action called")
        def load_model_action_for_ui(self): print("Controller: load_model_action_for_ui called by AppUI")

    root = tk.Tk()
    # root.withdraw() # Hide main window initially if loading screen is first
    controller = MockController()
    app_ui = AppUI(root, controller)
    
    # Test loading window
    # app_ui.show_loading_window(controller.load_model_action_for_ui)
    # To make loading window work standalone for test:
    # root.after(5000, lambda: app_ui.close_loading_window()) # Close after 5s
    # root.mainloop() # Need this if root was withdrawn and loading window is Toplevel
    
    # Test main UI (assuming loading is done)
    app_ui.setup_main_ui()
    app_ui.update_status("Main UI is active.")
    root.mainloop()
