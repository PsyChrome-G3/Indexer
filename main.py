import openpyxl
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.shared import Pt, Cm


def set_font(document, font_name):
    # Set the font of an element and its children in a Word document.
    styles = document.styles
    default_style = styles['Normal']
    default_font = default_style.font
    default_font.name = font_name
    default_font.size = Pt(11)  # Set the desired font size


def set_table_borders(table):
    border_xml = """
        <w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        </w:tblBorders>
    """
    table._element.xpath('//w:tblPr')[0].append(parse_xml(border_xml))


def create_tables_from_excel_rows(excel_file_path, sheet_name, word_file_path):
    # Load Excel workbook and select worksheet
    workbook = openpyxl.load_workbook(excel_file_path)
    worksheet = workbook[sheet_name]

    # Sort Excel rows alphabetically based on the first column
    sorted_rows = sorted(worksheet.iter_rows(min_row=2, values_only=True), key=lambda row: row[0])

    # Create a new Word document
    doc = Document()

    # Set the default font to Arial
    set_font(doc, 'Arial')

    # Calculate table width based on page size and margins
    section = doc.sections[0]
    page_width = section.page_width - section.left_margin - section.right_margin

    # Loop through each sorted row in Excel
    for row in sorted_rows:
        # Create a new table in Word with borders
        table = doc.add_table(rows=3, cols=2)
        table.style = 'Table Grid'

        # Set border properties for the table
        set_table_borders(table)

        # Set "Keep with next" option for the table
        tags = table._element.xpath('//w:tr[position() < last()]/w:tc/w:p')
        for tag in tags:
            ppr = tag.get_or_add_pPr()
            ppr.keepNext_val = True

        # Populate the table cells with the data from Excel
        entry1_cell = table.cell(0, 0)
        entry1_value = row[0] if row[0] else None  # Entry1
        if entry1_value is not None:
            entry1_cell.merge(table.cell(0, 1))
            entry1_cell.text = str(entry1_value)
            entry1_cell.paragraphs[0].runs[0].bold = True
            entry1_cell.paragraphs[0].runs[0].font.size = Pt(12)

        description_cell = table.cell(1, 0)
        description_value = row[3] if row[3] else None  # Description
        if description_value is not None:
            description_cell.merge(table.cell(1, 1))
            description_cell.text = str(description_value)

        book_cell = table.cell(2, 0)
        book_value = row[2] if row[2] else None  # Book
        if book_value is not None:
            book_cell.text = f"Book: {str(book_value)}"
            book_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            book_cell.paragraphs[0].paragraph_format.alignment = WD_ALIGN_VERTICAL.CENTER
            book_cell.paragraphs[0].runs[0].bold = True

        # Set border properties for the Page and Book cells
        page_cell = table.cell(2, 1)
        pages = row[1] if row[1] else None
        if pages is not None:
            pages = str(pages)
            if ',' in pages or '-' in pages:
                page_cell.text = f"Pages: {pages}"
            else:
                page_cell.text = f"Page: {pages}"
            page_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            page_cell.paragraphs[0].paragraph_format.alignment = WD_ALIGN_VERTICAL.CENTER
            page_cell.paragraphs[0].runs[0].bold = True

        # Get the available space on the current page
        available_space = section.page_height - section.top_margin - section.bottom_margin - section.header_distance - section.footer_distance

        # Add an empty paragraph after the table
        doc.add_paragraph()

    # Save the Word document
    doc.save(word_file_path)

    print("Your awesome index has been generated successfully!")


# Example usage
excel_file_path = "Index.xlsx"
sheet_name = "Sheet1"
word_file_path = "Awesome-Index.docx"

create_tables_from_excel_rows(excel_file_path, sheet_name, word_file_path)