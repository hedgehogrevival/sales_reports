from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

# Создаем новую рабочую книгу Excel
wb = Workbook()
ws = wb.active

# Добавляем тестовые данные
data = [
    ['Периоды', 'факт', 'БО'],
    ['2021', 40, 0],
    ['2022', 25, 0],
    ['2023', 35, 0],
    ['2024', 15, 0],
    ['2025', 40, 14]
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
categories = Reference(ws, min_col=1, min_row=2, max_row=6)
data = Reference(ws, min_col=2, max_col=3, min_row=1, max_row=6)

# Добавляем данные в диаграмму
chart.add_data(data, titles_from_data=True)
chart.set_categories(categories)

# Добавляем диаграмму на лист
ws.add_chart(chart, "A10")

# Сохраняем файл
wb.save("custom_stacked_chart.xlsx")