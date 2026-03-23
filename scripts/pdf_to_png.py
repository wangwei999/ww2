#!/usr/bin/env python3
"""
PDF to PNG converter using PyMuPDF
Usage: python pdf_to_png.py <input.pdf> <output_dir>
"""
import sys
import os
import fitz  # PyMuPDF

def pdf_to_png(pdf_path, output_dir):
    """Convert PDF pages to PNG images"""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Open the PDF
    doc = fitz.open(pdf_path)
    
    image_paths = []
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        
        # Set zoom for better quality (200 DPI equivalent)
        zoom = 200 / 72  # 72 is default DPI
        mat = fitz.Matrix(zoom, zoom)
        
        # Render page to image
        pix = page.get_pixmap(matrix=mat)
        
        # Save as PNG
        output_path = os.path.join(output_dir, f"page_{page_num + 1}.png")
        pix.save(output_path)
        image_paths.append(output_path)
        
        print(f"Created: {output_path}")
    
    doc.close()
    return image_paths

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python pdf_to_png.py <input.pdf> <output_dir>")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    output_dir = sys.argv[2]
    
    if not os.path.exists(pdf_path):
        print(f"Error: PDF file not found: {pdf_path}")
        sys.exit(1)
    
    try:
        paths = pdf_to_png(pdf_path, output_dir)
        print(f"Successfully converted {len(paths)} pages")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)
