import os
import img2pdf
from pypdf import PdfReader, PdfWriter
from pdf2docx import Converter
from pdf2image import convert_from_path
from docx2pdf import convert as docx_convert
from PIL import Image
import comtypes.client
from pptx import Presentation

class PDFEngine:
    @staticmethod
    def merge_pdfs(file_list, output_path):
        merger = PdfWriter()
        for pdf in file_list:
            merger.append(pdf)
        merger.write(output_path)
        merger.close()

    @staticmethod
    def split_pdf(input_path, output_folder):
        reader = PdfReader(input_path)
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        created_files = []
        for i, page in enumerate(reader.pages):
            writer = PdfWriter()
            writer.add_page(page)
            out_file = os.path.join(output_folder, f"{base_name}_page_{i+1}.pdf")
            with open(out_file, "wb") as f:
                writer.write(f)
            created_files.append(out_file)
        return created_files

    @staticmethod
    def rotate_pdf(input_path, output_path, degrees=90):
        reader = PdfReader(input_path)
        writer = PdfWriter()
        for page in reader.pages:
            page.rotate(degrees)
            writer.add_page(page)
        with open(output_path, "wb") as f:
            writer.write(f)

    @staticmethod
    def images_to_pdf(image_list, output_path):
        processed_images = []
        temp_created = []
        for img_path in image_list:
            try:
                img = Image.open(img_path)
                if img.mode == 'RGBA' or img.format != 'JPEG':
                    img = img.convert('RGB')
                    temp_path = img_path + ".temp.jpg"
                    img.save(temp_path, quality=95)
                    processed_images.append(temp_path)
                    temp_created.append(temp_path)
                else:
                    processed_images.append(img_path)
            except Exception as e:
                print(f"Error processing image {img_path}: {e}")

        if processed_images:
            with open(output_path, "wb") as f:
                f.write(img2pdf.convert(processed_images))
        
        for temp in temp_created:
            try: os.remove(temp)
            except: pass

    @staticmethod
    def pdf_to_images(input_path, output_folder, dpi=200, fmt="jpeg"):
        images = convert_from_path(input_path, dpi=dpi)
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        saved_files = []
        for i, img in enumerate(images):
            ext = fmt.lower()
            out_file = os.path.join(output_folder, f"{base_name}_page_{i+1:03d}.{ext}")
            img.save(out_file, fmt.upper())
            saved_files.append(out_file)
        return saved_files

    @staticmethod
    def pdf_to_word(input_path, output_path):
        cv = Converter(input_path)
        cv.convert(output_path, start=0, end=None)
        cv.close()

    @staticmethod
    def word_to_pdf(input_path, output_path):
        docx_convert(input_path, output_path)

    @staticmethod
    def pptx_to_pdf(input_path, output_path):
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1 
        try:
            deck = powerpoint.Presentations.Open(input_path, WithWindow=False)
            deck.SaveAs(output_path, 32) # 32 = ppSaveAsPDF
            deck.Close()
        except Exception as e:
            raise e
        finally:
            powerpoint.Quit()

    @staticmethod
    def pdf_to_pptx(input_path, output_path):
        # Convert PDF pages to images then place on slides
        temp_dir = os.path.dirname(output_path)
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        
        images = convert_from_path(input_path, dpi=150)
        prs = Presentation()
        blank_slide_layout = prs.slide_layouts[6] 
        
        if images:
            # Set slide size based on first page aspect ratio
            width, height = images[0].size
            # Pixels to EMU (approx 9525 per pixel at 96dpi)
            prs.slide_width = int(width * 9525) 
            prs.slide_height = int(height * 9525)

        for i, img in enumerate(images):
            temp_img_path = os.path.join(temp_dir, f"{base_name}_temp_slide_{i}.jpg")
            img.save(temp_img_path, 'JPEG')
            slide = prs.slides.add_slide(blank_slide_layout)
            slide.shapes.add_picture(temp_img_path, 0, 0, prs.slide_width, prs.slide_height)
            try: os.remove(temp_img_path)
            except: pass

        prs.save(output_path)

    @staticmethod
    def protect_pdf(input_path, output_path, password):
        reader = PdfReader(input_path)
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        writer.encrypt(password)
        with open(output_path, "wb") as f:
            writer.write(f)

    @staticmethod
    def compress_pdf(input_path, output_path):
        reader = PdfReader(input_path)
        writer = PdfWriter()
        for page in reader.pages:
            page.compress_content_streams()
            writer.add_page(page)
        with open(output_path, "wb") as f:
            writer.write(f)
