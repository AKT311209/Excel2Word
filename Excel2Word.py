import pandas as pd
import xlwings as xl
import docx as docu


def readexcel(sheetname, range_str):
    """ Reads data from an Excel sheet within a specified range. """
    wb = xl.Book('Test.xlsx')
    sheet = wb.sheets[sheetname]
    df = pd.DataFrame(sheet.range(range_str).value)
    return df


def create_table_in_docx(data, filename="Test.docx"):
    doc = docu.Document(filename)

    # Check if a table exists
    if doc.tables:
        table = doc.tables[0]
        # Modify the table data while preserving formatting
        for i, row in enumerate(data.itertuples(index=False)):
            for j, cell_value in enumerate(row):
                # Get the existing paragraph in the cell
                paragraph = table.cell(i, j).paragraphs[0]
                # Check if the cell is empty
                if not paragraph.runs:
                    # Add a new run with the cell value and copy formatting
                    new_run = paragraph.add_run(str(cell_value))
                    # Copy formatting from the previous cell (if available)
                    if i > 0 and j > 0:
                        previous_cell_paragraph = table.cell(
                            i - 1, j - 1).paragraphs[0]
                        if previous_cell_paragraph.runs:
                            new_run.font.name = previous_cell_paragraph.runs[0].font.name
                            new_run.font.size = previous_cell_paragraph.runs[0].font.size
                            new_run.font.bold = previous_cell_paragraph.runs[0].font.bold
                            new_run.font.italic = previous_cell_paragraph.runs[0].font.italic
                            # Add more formatting properties as needed
                else:
                    # Replace the existing text with the new text
                    paragraph.runs[0].text = str(
                        cell_value)  # Update the first run
    else:
        # Create a new table
        table = doc.add_table(rows=len(data) + 1, cols=len(data.columns))
        # Populate the new table
        for i, row in enumerate(data.itertuples(index=False)):
            for j, cell_value in enumerate(row):
                table.cell(i, j).text = str(cell_value)
    doc.save(filename)


if __name__ == "__main__":
    sheet_name = input("Enter the sheet name: ")
    range_str = input("Enter the range (e.g., A1:C5): ")
    file_name = input("Enter the output file name: ")
    if file_name:
        file_name = file_name + ".docx"

    data = readexcel(sheet_name, range_str)
    create_table_in_docx(data, file_name)

    print("Table created successfully!")
