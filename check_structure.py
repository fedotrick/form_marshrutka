from openpyxl import load_workbook

wb = load_workbook('plavka.xlsx')
sheet = wb.active

print("Заголовки столбцов:")
for idx, cell in enumerate(sheet[1], 1):
    print(f"Столбец {idx}: {cell.value}")

# Показать первую строку с данными
print("\nПример данных (первая строка):")
for idx, cell in enumerate(sheet[2], 1):
    print(f"Столбец {idx}: {cell.value}")

wb.close()
