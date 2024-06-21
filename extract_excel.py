import openpyxl

def fill_values(file_path):
    """
    Fills the values of fields from the given Excel file.
    """
    fields = [
        "Total revenue from operations",
        "Changes in inventories of finished goods, work-in-progress and stock-in-trade",
        "Total other income",
        "Total income",
        "Changes in inventories of finished goods, work-in-progress and stock-in-trade",
        ""
    ]

    field_values = {}

    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    for row in ws.iter_rows(min_row=1, values_only=True):
        if row[0] is not None:
            field_name = row[0].strip().lower().replace(" ", "_")
            if field_name in fields:
                field_values[field_name] = list(row[1:])

    return field_values