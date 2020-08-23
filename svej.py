import win32com.client
import ctypes
from win32api import GetSystemMetrics
import os
import colorama
from colorama import Fore, Back, Style

kernel32 = ctypes.WinDLL('kernel32')
user32 = ctypes.WinDLL('user32')

SW_MAXIMIZE = 3

hWnd = kernel32.GetConsoleWindow()
user32.ShowWindow(hWnd, SW_MAXIMIZE)
os.system("mode con cols=500 lines=200")


print("Width =", GetSystemMetrics(0))
print("Height =", GetSystemMetrics(1))

# if GetSystemMetrics(0) >= 1366 and GetSystemMetrics(1) >= 768:
#     Window.size = (GetSystemMetrics(0), GetSystemMetrics(1))
#     Window.fullscreen = True
# else: 
#     Window.size = (1366, 768)

colorama.init()
print(Fore.YELLOW + 'Работа скрипта начата')

#Запускает приложение Эксель
Excel = win32com.client.Dispatch('Excel.Application')
#Позволяет редактировать файл
Excel.Visible = True
#Открывает книга1
wb = Excel.Workbooks.Open(u'C:\\Книга1.xls')
# Открывает 2 листа в книге
sheet1 = wb.Sheets(1)
sheet2 = wb.Sheets(2)
# sheet3 = wb.Sheets(3)

#
end_num1 = sheet1.UsedRange.Rows.Count
end_num2 = sheet2.UsedRange.Rows.Count

vals1 = [r1[0].value for r1 in sheet1.Range("B5:B" + str(end_num1))]
# cell1 = [c1[0].value for c1 in sheet1.Range("D5:D"+ str(end_num1))]
cell11 = [c11[0].value for c11 in sheet1.Range("E5:E"+ str(end_num1))]

vals2 = [r2[0].value for r2 in sheet2.Range("A8:A"+ str(end_num2))]
cell2 = [c2[0].value for c2 in sheet2.Range("E8:E"+ str(end_num2))]

# sheet1.Cells(5, 5).Value = 55
# sheet1.Cells(6, 5).Value = 55
# sheet1.Cells(7, 5).Value = 55
# print(cell2[end_num2 - 10])
for i  in range(end_num1 - 4):
	# print(vals1[i])
	for j in range(end_num2 - 14):
		# print(vals2[j][1:-6])
		if vals1[i] == vals2[j][1:-6]:
			print()
			print( Back.WHITE + Fore.GREEN + 'Найдено значение: ' + Back.WHITE + str(vals1[i]))
			# if str(cell2[j+1]) == '':
			# 	pole = cell2[j+1]
			# 	pole = 'Пустое поле'
			# 	sheet1.Cells(i+5, 5).Value = pole
			sheet1.Cells(i+5, 5).Value = cell2[j+1]
			print(Fore.RED + 'Status: ====>' + '\t' + Back.WHITE + str(cell2[j+1]) + Fore.YELLOW + ' значение записано')
sheet1.Cells(i+5, 5).Value = cell2[end_num2 - 10]
# i = 5
# for i in range(end_num1 - 4):
# 	sheet1.Cells(i, 5).Value = 55
# 	# print(vals1[i] + '    ' + str(cell1[i]))

# 	print(cell1[i])
# 	# print(cell11[i])

# j = 0
# for val in vals1:
# 	j = j + 1
# 	print(str(j) + ' ' + val) 

# vals2 = [r2[0].value for r2 in sheet2.Range("A8:A"+ str(end_num2))]
# cell2 = [c2[0].value for c2 in sheet2.Range("E8:E"+ str(end_num2))]


# print(vals2[0][1:-6])
# print(vals2[2][1:-6])
# print(cell2[1])
# print()
# # print(vals3[0])
# # print(vals3[3])
# # print(vals2[20][1:-6])
# # print(cell2[21])

# if vals4[1] == vals2[20][1:-6]:
# 	print('True')
# 	print(vals3[1])
# 	print(cell2[21])


# #Вывод значения ячейки 1 1
# var1 = sheet.Cells(first_row, 1).Value
# print(type(var1))


# #Запись числа в ячейку 3 2
# sheet.Cells(3, 2).Value = 55
speaker = win32com.client.Dispatch("SAPI.SpVoice")
# speaker = win32com.client.Dispatch("Speech.SpVoice")
speaker.Speak("The script completed successfully")

print()
print()
print()

print(Fore.BLUE + 'Работа скрипта выполнена успешно')

# #Выводит в консоль ячейку 1 1
# print(var1)

wb.Save()
# wb.Close()
# Excel.Quit()
