import openpyxl

def shift_and_delete_rows(file_path):
    # Load the workbook and select the active worksheet
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    # Unmerge all cells in the first row
    for merged_cell in list(ws.merged_cells.ranges):
        if merged_cell.min_row == 1:
            ws.unmerge_cells(str(merged_cell))

    # Explicitly unmerge the cells G2:K2
    ws.unmerge_cells('G2:K2')

    # Shift rows down by 1
    max_row = ws.max_row
    max_col = ws.max_column

    # Copy each row to the row below it
    for row in range(max_row, 1, -1):
        for col in range(1, max_col + 1):
            ws.cell(row=row+1, column=col).value = ws.cell(row=row, column=col).value

    # Delete the first two rows
    ws.delete_rows(1, 2)

    # Move values from G2:K2 to G1:K1
    for col in range(7, 12):  # G=7, K=11
        ws.cell(row=1, column=col).value = ws.cell(row=2, column=col).value

    # Clear all values in G2:K2
    for col in range(7, 12):  # G=7, K=11
        ws.cell(row=2, column=col).value = None

    # Change the value of B1 to "Courses"
    ws['B1'] = "Courses"

    # Save the modified workbook
    wb.save(file_path)

    print(f"File saved as {file_path}")