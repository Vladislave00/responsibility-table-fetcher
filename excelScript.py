import openpyxl
import openpyxl.styles
import database


# Метод для форматирования границ таблицы
def set_border(ws, cell_range, need_to_thick, need_to_thick_up, need_to_thick_down):
    thin = openpyxl.styles.Side(border_style="thin", color="000000")
    thick = openpyxl.styles.Side(border_style="thick", color="000000")
    for row in ws[cell_range]:
        for cell in row:
            if cell == row[len(row) - 1]:
                cell.border = openpyxl.styles.Border(top=thin, left=thin, right=thick, bottom=thin)
                cell.fill = openpyxl.styles.PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
            else:
                cell.border = openpyxl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
        if len(need_to_thick) != 0:
            if row == ws[need_to_thick[0]]:
                for cell in row:
                    cell.border = openpyxl.styles.Border(top=thick, left=thick, right=thick, bottom=thick)
                need_to_thick.pop(0)
        if len(need_to_thick_up) != 0:
            if row == ws[need_to_thick_up[0]]:
                for cell in row:
                    if cell == row[len(row) - 1]:
                        cell.border = openpyxl.styles.Border(top=thick, left=thin, right=thick, bottom=thin)
                    else:
                        cell.border = openpyxl.styles.Border(top=thick, left=thin, right=thin, bottom=thin)
                need_to_thick_up.pop(0)
        if len(need_to_thick_down) != 0:
            if row == ws[need_to_thick_down[0]]:
                for cell in row:
                    if cell == row[len(row) - 1]:
                        cell.border = openpyxl.styles.Border(top=thin, left=thin, right=thick, bottom=thick)
                    else:
                        cell.border = openpyxl.styles.Border(top=thin, left=thin, right=thin, bottom=thick)

                need_to_thick_down.pop(0)


# Метод для формирования Excel таблицы
def make_excel():
    # Получаем метрики из базы данных
    data = database.Database.get_data()
    # Создаем книгу и заполняем ее
    wb = openpyxl.Workbook()

    # добавляем новый лист
    wb.create_sheet(title='Зоны ответственности', index=0)

    # получаем лист, с которым будем работать
    sheet = wb['Зоны ответственности']
    ws = wb.active

    # формирование шапки таблицы
    header_data = ["Активность", "Подназвание активности", "Номер", "Буква", "Индикатор", "Имя сотрудника", "Должность"]
    ws.append(header_data)


    # Заполнение строк данными из базы данных
    for row in data:
        ws.append(row)

    # Выставление ширины колонок
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 40
    ws.column_dimensions['F'].width = 20
    ws.column_dimensions['G'].width = 20

    # Сохранение файла
    wb.save('Зоны ответственности.xlsx')
    print("Файл успешно сохранен")


make_excel()
