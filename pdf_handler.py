import fitz  # PyMuPDF
from PIL import Image
import numpy as np
from paddleocr import PaddleOCR # For OCR capabilities

class PDFHandler:
    def __init__(self, use_gpu=True, lang='ch'):
        """
        Initializes the PDFHandler, including the PaddleOCR engine.
        Args:
            use_gpu (bool): Whether PaddleOCR should attempt to use GPU.
            lang (str): Language for PaddleOCR.
        """
        # Store args for potential re-initialization or info
        self.use_gpu = use_gpu
        self.lang = lang
        self.ocr_engine_initialized = False
        try:
            # print(f"Initializing PaddleOCR with lang={lang}, use_gpu={use_gpu}") # Debug
            self.ocr = PaddleOCR(use_angle_cls=True, lang=self.lang, use_gpu=self.use_gpu, show_log=False)
            self.ocr_engine_initialized = True
            # print("PaddleOCR engine initialized successfully.") # Debug
        except Exception as e:
            # print(f"Error initializing PaddleOCR in PDFHandler: {e}") # Debug
            # This error should be propagated or handled in a way the app can inform the user.
            self.ocr = None 
            raise RuntimeError(f"Failed to initialize PaddleOCR in PDFHandler: {e}")

    def _pdf_to_images(self, pdf_path, update_status_callback=None):
        """
        Converts each page of a PDF file into a list of images (NumPy arrays).
        Internal method, typically used for scanned PDFs requiring OCR.
        """
        try:
            doc = fitz.open(pdf_path)
        except Exception as e:
            if update_status_callback:
                update_status_callback(f"Error opening PDF: {e}")
            # print(f"Error opening PDF {pdf_path}: {e}") # Debug
            return [] # Return empty list on error

        images = []
        if update_status_callback:
            update_status_callback(f"Converting PDF to images (Total pages: {len(doc)})...")
        
        for i, page in enumerate(doc):
            if update_status_callback:
                update_status_callback(f"Converting page {i+1}/{len(doc)} to image...")
            try:
                pix = page.get_pixmap()
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                images.append(np.array(img))
            except Exception as e_page:
                if update_status_callback:
                    update_status_callback(f"Error converting page {i+1}: {e_page}")
                # print(f"Error converting page {i+1} of {pdf_path}: {e_page}") # Debug
                continue # Skip problematic page
        doc.close()
        return images

    def _is_scanned_pdf(self, pdf_path, update_status_callback=None):
        """
        Checks if a PDF is likely a scanned document (image-based).
        Internal method.
        """
        text_content_len = 0
        try:
            doc = fitz.open(pdf_path)
            # Check a few pages for text content, not necessarily all for performance.
            num_pages_to_check = min(len(doc), 3) # Check up to 3 pages
            if num_pages_to_check == 0 and update_status_callback:
                update_status_callback("Warning: PDF has 0 pages.")
            
            for i in range(num_pages_to_check):
                page = doc.load_page(i)
                text_content_len += len(page.get_text("text").strip())
                if text_content_len >= 100: # If substantial text found early, assume not scanned
                    doc.close()
                    return False
            doc.close()
        except Exception as e:
            if update_status_callback:
                update_status_callback(f"Error checking PDF scan status: {e}")
            # print(f"Error checking PDF scan status for {pdf_path}: {e}") # Debug
            return True # Default to scanned if error occurs during check

        return text_content_len < 100 # Heuristic: less than 100 chars means likely scanned

    def extract_text_from_pdf(self, pdf_path, update_status_callback=None):
        """
        Extracts text from a PDF, handling both scanned and text-based documents.
        Args:
            pdf_path (str): Path to the PDF file.
            update_status_callback (function, optional): Callback to update UI status.
        Returns:
            list: A list of strings, where each string is a line of extracted text.
                  Returns empty list if OCR engine failed to initialize or no text found.
        """
        if not self.ocr_engine_initialized or not self.ocr:
            if update_status_callback:
                update_status_callback("Critical Error: OCR engine not available.")
            # print("Critical Error: PDFHandler's OCR engine not initialized or failed.") # Debug
            # This case should ideally be prevented by handling the RuntimeError in __init__
            raise RuntimeError("OCR engine is not available in PDFHandler.")

        text_content_lines = []
        try:
            if self._is_scanned_pdf(pdf_path, update_status_callback):
                if update_status_callback:
                    update_status_callback("检测到扫描PDF，正在进行OCR识别...")
                
                images = self._pdf_to_images(pdf_path, update_status_callback)
                total_images = len(images)
                if total_images == 0 and update_status_callback:
                    update_status_callback("未能从PDF转换任何图像进行OCR。")
                    return []

                for i, img in enumerate(images):
                    if update_status_callback:
                        update_status_callback(f"正在OCR第 {i+1}/{total_images} 页...")
                    
                    ocr_result = self.ocr.ocr(img, cls=True) 
                    if ocr_result and ocr_result[0] is not None:
                         for line_info in ocr_result[0]: 
                            text_content_lines.append(line_info[1][0])
            else:
                if update_status_callback:
                    update_status_callback("检测到可复制PDF，直接提取文本...")
                
                doc = fitz.open(pdf_path)
                total_pages = len(doc)
                if total_pages == 0 and update_status_callback:
                    update_status_callback("PDF文件没有页面。")
                    doc.close()
                    return []

                for i, page in enumerate(doc):
                    if update_status_callback:
                        update_status_callback(f"正在提取第 {i+1}/{total_pages} 页文本...")
                    text_content_lines.extend(page.get_text("text").splitlines())
                doc.close()
        except Exception as e:
            if update_status_callback:
                update_status_callback(f"提取文本时发生错误: {e}")
            # print(f"Error during text extraction from {pdf_path}: {e}") # Debug
            # Depending on severity, might re-raise or return empty/partial list
            return [] # Return empty on error for now
        
        if not text_content_lines and update_status_callback:
            update_status_callback("未能从PDF中提取任何文本内容。")
        elif text_content_lines and update_status_callback:
            update_status_callback("PDF文本提取完成。")
            
        return text_content_lines
