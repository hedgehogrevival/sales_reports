from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

# Создаем новую книгу и активный лист
wb = Workbook()
ws = wb.active

# Добавляем тестовые данные
data = [
    ["Категория", "Значения1", "Значения2"],
    ["A", 10, 15],
    ["B", 20, 25],
    ["C", 30, 35],
    ["D", 40, 45],
]
for row in data:
    ws.append(row)

# Создаем диаграмму
chart = BarChart()
chart.title = "Пример столбчатой диаграммы"
chart.x_axis.title = "Категории"
chart.y_axis.title = "Значения1"
chart.y_axis.title = "Значения2"

# Определяем данные для диаграммы
data_ref = Reference(ws, min_col=2, min_row=1, max_row=5, max_col=3)
categories_ref = Reference(ws, min_col=1, min_row=2, max_row=5)

# Добавляем данные в диаграмму
chart.add_data(data_ref, titles_from_data=True)
chart.set_categories(categories_ref)

# Включаем подписи данных
chart.dataLabels = DataLabelList(showVal=True, showCatName=False, showSerName=False, showLegendKey=False)
#chart.dataLabels.showVal = True  # Показывать значения

# Размещаем диаграмму на листе
ws.add_chart(chart, "E5")

# Сохраняем книгу
wb.save("bar_chart_with_labels.xlsx")