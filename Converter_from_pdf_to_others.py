﻿from pdf2docx import Converter
from pptx import Presentation
from pdf2image import convert_from_path
import pdfplumber
import pandas as pd


def pdf_to_docx(pdf_file, output_file, page=None, start_page=0, end_page=None):
    cv = Converter(pdf_file)
    try:
        if page is not None:   
            cv.convert(output_file, pages=page)       
        else:       
            cv.convert(output_file, start=start_page, end=end_page)
        print(f'Done: {output_file}')
        cv.close()
    except:
        cv.close()


def pdf_to_xlsx(pdf_file, xlsx_file, page_range=None):
    xlsx_file += ".xlsx"
    if page_range is not None:
        page_range = [int(page) + 1 for page in page_range]
    with pdfplumber.open(pdf_file) as pdf:
        all_data = []
        for page_number, page in enumerate(pdf.pages):
            if page_range is not None and page_number + 1 not in page_range:
                continue
            tables = page.extract_tables()
            for table_number, table in enumerate(tables):
                if not table:
                    continue
                
                df = pd.DataFrame(table)
                
                if df.empty:
                    continue

                all_data.append(df)

        if page_range is None and all_data:
            with pd.ExcelWriter(xlsx_file, engine='xlsxwriter') as writer:
                for i, df in enumerate(all_data):
                    df.to_excel(writer, sheet_name=f"Table_{i+1}", index=False, header=False)
            print(f'Done: {xlsx_file}')
            
        elif page_range is not None and all_data:
            with pd.ExcelWriter(xlsx_file, engine='xlsxwriter') as writer:
                for i, df in enumerate(all_data):
                    df.to_excel(writer, sheet_name=f"Page_{page_range[i]}_Table_{i+1}", index=False, header=False)
            print(f'Done: {xlsx_file}')
        else:
            print("No data found.")

    

def pdf_to_pptx(pdf_file, pptx_file, page):
    pass


def pdf_to_images(pdf_file, image_dir, start_page, end_page):
    images = convert_from_path(pdf_file, first_page=start_page, last_page=end_page)
    for i, image in enumerate(images):
        image.save(f"{image_dir}/page_{i+1}.jpg", "JPEG")
        

def convert_file(pdf_file, output_file, choice_type, page, start_page, end_page, doc_type):
    if choice_type == "1":
        pdf_to_docx(pdf_file, output_file + "." + doc_type, page, start_page, end_page)
    elif choice_type == "2":
        pdf_to_xlsx(pdf_file, output_file, page)
        print(page)
             

def choice_page(number_of_pages):
    start_page = None
    end_page = None
    page = None
    if number_of_pages == "1":
        start_page = str(int(input("Start: ")) - 1)
        end_page = str(int(input("End: ")) - 1)
        return start_page, end_page, None
    elif number_of_pages == "2":
        page = [str(int(input("Page: ")) - 1)] 
        return None, None, page
    else:
        return None, 0, None
    

def output_file_preparation(pdf_file, output_file):
    parts = pdf_file.split('/') 
    file_name = parts[-1].split(".")
    if output_file == "":
        output_file += file_name[0]
    else:
        output_file += "/" + file_name[0]
        
    return output_file


def main():
    pdf_file = input("Enter path and file name: ")
    choice_type = input("Choose an option:\n1. Convert to DOCX/DOC\n2. Convert to XLSX\n3. Convert to PPTX\n4. Convert to Images\n")
    if choice_type == "1":
        doc_type = input("Enter doc/docx: ")
    number_of_pages = input("Choose an option:\n1. Page range\n2. Only one page\n3. All file\n")
    doc_type = None
    
    start_page, end_page, page = choice_page(number_of_pages)
        
    output_file = input("Enter the path where you want to save the file: ") 
    output_file = output_file_preparation(pdf_file, output_file)
    
    convert_file(pdf_file, output_file, choice_type, page,  start_page, end_page, doc_type)
    

if __name__ == "__main__":
    main()