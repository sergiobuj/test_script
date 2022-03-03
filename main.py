from csv import DictReader
import openpyxl as opxl

MAPPING = {
    "banco":"B8",
    "beneficiario":"B5",
    "ccnit":"H5",
    "concepto":"B6",
    "dd":"B2",
    "mm": "C2",
    "numero de cuenta":"G8",
    "orden de pago": "I1",
    "valor":"B4",
    "yyyy": "D2",
}

XLSX_TEMPLATE = "template.xlsx"
CSV_FILE = "data.csv"
XLSX_OUTPUT = "out.xlsx"

def run():
    wb = opxl.load_workbook(XLSX_TEMPLATE)

    with open(CSV_FILE) as csvfile:
        reader = DictReader(csvfile)
        ws = wb.active # Get the active worksheet (the first one named "Ejemplo")
        for row in reader:
            ws = wb.copy_worksheet(ws)
            ws.title = f"Orden de pago {row['orden de pago']}"
            for k in row.keys():
                if k in MAPPING:
                    ws[MAPPING[k]] = row[k]
        del wb["Ejemplo"]

    for ws in wb.worksheets:
        print(ws.title)

    wb.save(XLSX_OUTPUT)
    print("Done")

if __name__ == '__main__':
    run()
