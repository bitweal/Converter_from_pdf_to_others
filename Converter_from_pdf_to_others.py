from pdf2docx import Converter
import tabula
from pptx import Presentation
from pdf2image import convert_from_path


def pdf_to_docx(pdf_file, output_file, page=None, start_page=0, end_page=None):
    cv = Converter(pdf_file)
    if page is not None:   
        cv.convert(output_file, pages=page)       
    else:       
        cv.convert(output_file, start=start_page, end=end_page)
    cv.close()


def pdf_to_xlsx(pdf_file, xlsx_file, page):
    tabula.convert_into(pdf_file, xlsx_file, output_format="xlsx", pages=page)


def pdf_to_pptx(pdf_file, pptx_file, page):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    title = slide.shapes.title
    content = slide.placeholders[1]
    content.text = tabula.read_pdf(pdf_file, pages=page).values[0][0]
    prs.save(pptx_file)


def pdf_to_images(pdf_file, image_dir, start_page, end_page):
    images = convert_from_path(pdf_file, first_page=start_page, last_page=end_page)
    for i, image in enumerate(images):
        image.save(f"{image_dir}/page_{i+1}.jpg", "JPEG")
        

def convert_file(pdf_file, output_file, choice_type, page, start_page, end_page, doc_type):
    if choice_type == "1":
        pdf_to_docx(pdf_file, output_file + "." + doc_type, page, start_page, end_page)
             

def main():
    pdf_file = input("Enter path and file name: ")
    choice_type = input("Choose an option:\n1. Convert to DOCX/DOC\n2. Convert to XLSX\n3. Convert to PPTX\n4. Convert to Images\n")
    if choice_type == "1":
        doc_type = input("Enter doc/docx: ")
    number_of_pages = input("Choose an option:\n1. Page range\n2. Only one page\n3. All file\n")
    start_page = None
    end_page = None
    page = None
    
    if number_of_pages == "1":
        start_page = str(int(input("Start: ")) - 1)
        end_page = str(int(input("End: ")) - 1)
    elif number_of_pages == "2":
        page = [str(int(input("Page: ")) - 1)] 
    else:
        start_page = 0
        
    output_file = input("Enter the path where you want to save the file: ")  
    parts = pdf_file.split('/') 
    file_name = parts[-1].split(".")
    if output_file == "":
        output_file += file_name[0]
    else:
        output_file += "/" + file_name[0]
    convert_file(pdf_file, output_file, choice_type, page,  start_page, end_page, doc_type)
    

if __name__ == "__main__":
    main()