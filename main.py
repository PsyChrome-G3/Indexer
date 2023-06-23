import openpyxl
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.shared import Pt


def create_tables_from_excel_rows(excel_file_path, sheet_name, word_file_path):
    # Load Excel workbook and select worksheet
    workbook = openpyxl.load_workbook(excel_file_path)
    worksheet = workbook[sheet_name]

    # Create a new Word document
    doc = Document()

    # Calculate table width based on page size and margins
    section = doc.sections[0]
    page_width = section.page_width - section.left_margin - section.right_margin

    # Loop through each row in Excel
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        # Create a new table in Word with borders
        table = doc.add_table(rows=3, cols=2)
        table.style = 'Table Grid'

        # Set border properties for the table
        border_xml = """
            <w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            </w:tblBorders>
        """
        table._element.xpath('//w:tblPr')[0].append(parse_xml(border_xml))

        # # Adjust column widths as percentages
        # column_widths = [int(0.5 * page_width), int(0.5 * page_width)]
        # for colIndex, width in enumerate(column_widths):
        #     table.columns[colIndex].width = width

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