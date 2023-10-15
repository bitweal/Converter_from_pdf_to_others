from pdf2docx import Converter
import tabula
from pptx import Presentation
from pdf2image import convert_from_path


def pdf_to_docx(pdf_file, docx_file):
    cv = Converter(pdf_file)
    cv.convert(docx_file, start=0, end=None)
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


pdf_file = input("Enter the PDF file name: ")
docx_file = 'output.docx'
xlsx_file = 'output.xlsx'
pptx_file = 'output.pptx'
image_dir = 'images'

# Choose what to convert
print("1. Convert the entire file")
print("2. Convert a single page")
print("3. Convert a page range")
choice = input("Choose an option (1/2/3): ")

if choice == "1":
    pdf_to_docx(pdf_file, docx_file)
    pdf_to_xlsx(pdf_file, xlsx_file, "all")
    pdf_to_pptx(pdf_file, pptx_file, "all")
    pdf_to_images(pdf_file, image_dir, 1, None)
elif choice == "2":
    page = input("Enter the page number: ")
    pdf_to_xlsx(pdf_file, xlsx_file, page)
    pdf_to_pptx(pdf_file, pptx_file, page)
    pdf_to_images(pdf_file, image_dir, int(page), int(page))
elif choice == "3":
    start_page = int(input("Enter the starting page: "))
    end_page = int(input("Enter the ending page: "))
    pdf_to_xlsx(pdf_file, xlsx_file, f"{start_page}-{end_page}")
    pdf_to_pptx(pdf_file, pptx_file, f"{start_page}-{end_page}")
    pdf_to_images(pdf_file, image_dir, start_page, end_page)