import openpyxl
import database
#Получаем метрики из базы данных
data = database.Database.get_metrics()
#Создаем книгу и заполняем ее
wb = openpyxl.Workbook()

# добавляем новый лист
wb.create_sheet(title = 'Метрики', index = 0)

# получаем лист, с которым будем работать
sheet = wb['Метрики']
ws=wb.active 
for row in range(len(data)):
    for col in range(1, len(data[row])):
        cell = sheet.cell(row = row+1, column = col)
        #Для пропуска записи в недоступную ячейку после слияния
        try:
            cell.value = data[row][col]
        except Exception as err:
            continue
    #слияние ячеек с номерам критериев
    if row+1 < len(data):
        if data[row][1] == data[row+1][1]:
                ws.merge_cells(start_row=row+1, start_column=1, end_row=row+2, end_column=1)  


wb.save('example.xlsx')



