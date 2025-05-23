import fitz  # PyMuPDF
from PIL import Image
import numpy as np

class PDFHandler:
    # Converts each page of a PDF file into a list of images (NumPy arrays).
    # This is used when OCR is required for scanned PDFs.
    def pdf_to_images(self, pdf_path):
        # Parameters:
        #   pdf_path (str): The path to the PDF file.
        # Returns:
        #   list: A list of NumPy arrays, where each array represents an image of a PDF page.
        doc = fitz.open(pdf_path) # Open the PDF file using PyMuPDF (fitz).
        images = [] # Initialize an empty list to store images.
        for page in doc: # Iterate through each page in the PDF.
            pix = page.get_pixmap() # Render the page as a pixmap (image).
            # Convert the pixmap to a PIL Image object.
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            images.append(np.array(img)) # Convert PIL Image to NumPy array and add to list.
        doc.close() # Close the PDF document.
        return images

    # Checks if a PDF is likely a scanned document (image-based) rather than text-based.
    # It does this by trying to extract text and checking if the amount of text is very small.
    def is_scanned_pdf(self, pdf_path):
        # Parameters:
        #   pdf_path (str): The path to the PDF file.
        # Returns:
        #   bool: True if the PDF is likely scanned, False otherwise.
        # Docstring in Chinese: """检查PDF是否为扫描件""" (Checks if PDF is a scanned document)
        doc = fitz.open(pdf_path) # Open the PDF file.
        text_content = "" # Initialize an empty string to store extracted text.
        for page in doc: # Iterate through each page.
            text_content += page.get_text() # Append extracted text from the page.
        doc.close() # Close the PDF.
        # If the total length of stripped text content from all pages is less than 100 characters,
        # it's considered a scanned PDF. This is a heuristic.
        return len(text_content.strip()) < 100
