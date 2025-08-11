from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

# Создаем новую рабочую книгу Excel
wb = Workbook()
ws = wb.active

# Добавляем тестовые данные
data = [
    ['Категория', '2020', '2021', '2022'],
    ['Продукт A', 40, 30, 25],
    ['Продукт B', 30, 25, 30],
    ['Продукт C', 20, 35, 30],
    ['Продукт D', 10, 10, 15],
]

for row in data:
    ws.append(row)

# Создаем диаграмму с накоплением
chart = BarChart()
chart.type = "col"
chart.style = 10
chart.grouping = "stacked"
chart.overlap = 100
chart.title = "Продажи по годам (с накоплением)"
chart.y_axis.title = 'Объем продаж'
chart.x_axis.title = 'Категории'

# Определяем диапазоны данных
categories = Reference(ws, min_col=1, min_row=2, max_row=5)
data = Reference(ws, min_col=2, max_col=4, min_row=1, max_row=5)

# Добавляем данные в диаграмму
chart.add_data(data, titles_from_data=True)
chart.set_categories(categories)

# Добавляем диаграмму на лист
ws.add_chart(chart, "A10")

# Сохраняем файл
wb.save("stacked_chart.xlsx")