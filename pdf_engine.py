import os
import img2pdf
from pypdf import PdfReader, PdfWriter
from pdf2docx import Converter
from pdf2image import convert_from_path
from docx2pdf import convert as docx_convert
from PIL import Image
import comtypes.client
from pptx import Presentation
import pikepdf

class PDFEngine:
    @staticmethod
    def merge_pdfs(file_list, output_path):
        merger = PdfWriter()
        for pdf in file_list:
            merger.append(pdf)
        merger.write(output_path)
        merger.close()

    @staticmethod
    def split_pdf(input_path, output_folder, mode="all", page_range=None):
        reader = PdfReader(input_path)
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        total_pages = len(reader.pages)
        
        selected_indices = []
        if page_range:
            try:
                parts = [p.strip() for p in page_range.split(',')]
                for p in parts:
                    if '-' in p:
                        start, end = map(int, p.split('-'))
                        selected_indices.extend(range(start-1, end))
                    else:
                        selected_indices.append(int(p)-1)
                selected_indices = [i for i in selected_indices if 0 <= i < total_pages]
            except:
                pass 
        else:
            selected_indices = list(range(total_pages))

        if mode == "all":
            for i in selected_indices:
                writer = PdfWriter()
                writer.add_page(reader.pages[i])
                out_file = os.path.join(output_folder, f"{base_name}_page_{i+1}.pdf")
                with open(out_file, "wb") as f:
                    writer.write(f)
        
        elif mode == "extract":
            writer = PdfWriter()
            for i in selected_indices:
                writer.add_page(reader.pages[i])
            out_file = os.path.join(output_folder, f"{base_name}_extracted.pdf")
            with open(out_file, "wb") as f:
                writer.write(f)

    @staticmethod
    def reorder_save_pdf(input_path, output_path, page_order_data):
        reader = PdfReader(input_path)
        writer = PdfWriter()
        for item in page_order_data:
            idx = item['original_index']
            rotation = item.get('rotation', 0)
            if 0 <= idx < len(reader.pages):
                page = reader.pages[idx]
                if rotation != 0:
                    page.rotate(rotation)
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
    def compress_pdf(input_path, output_path, level="medium"):
        if level == "low":
            reader = PdfReader(input_path)
            writer = PdfWriter()
            for page in reader.pages:
                page.compress_content_streams()
                writer.add_page(page)
            writer.add_metadata(reader.metadata)
            with open(output_path, "wb") as f:
                writer.write(f)
        elif level == "medium":
            try:
                with pikepdf.open(input_path) as pdf:
                    pdf.save(output_path, compress_streams=True, object_stream_mode=pikepdf.ObjectStreamMode.generate)
            except Exception as e:
                raise Exception(f"Pikepdf failed: {e}")
        elif level == "extreme":
            import tempfile, shutil
            temp_dir = tempfile.mkdtemp()
            try:
                imgs = PDFEngine.pdf_to_images(input_path, temp_dir, dpi=130, fmt="jpeg")
                with open(output_path, "wb") as f:
                    f.write(img2pdf.convert(imgs))
            finally:
                shutil.rmtree(temp_dir)

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
            deck.SaveAs(output_path, 32)
            deck.Close()
        except Exception as e:
            raise e
        finally:
            powerpoint.Quit()

    @staticmethod
    def pdf_to_pptx(input_path, output_path):
        temp_dir = os.path.dirname(output_path)
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        images = convert_from_path(input_path, dpi=150)
        prs = Presentation()
        blank_slide_layout = prs.slide_layouts[6] 
        if images:
            width, height = images[0].size
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
    def protect_pdf(input_path, output_path, password, algorithm="AES-256"):
        """
        Algorithms: 'AES-256', 'AES-128', 'RC4-128'
        """
        reader = PdfReader(input_path)
        writer = PdfWriter()
        writer.append_pages_from_reader(reader)
        
        # pypdf encryption logic
        # For compatibility, we map our UI names to pypdf expected values if needed, 
        # but pypdf handles standard strings well in modern versions.
        writer.encrypt(user_password=password, algorithm=algorithm)
        
        with open(output_path, "wb") as f:
            writer.write(f)

    @staticmethod
    def unlock_pdf(input_path, output_path, password):
        """
        Decrypts a PDF and saves it to output_path.
        """
        reader = PdfReader(input_path)
        
        if reader.is_encrypted:
            success = reader.decrypt(password)
            if not success:
                raise Exception("Incorrect Password")
        
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
            
        with open(output_path, "wb") as f:
            writer.write(f)
        
        return output_path
