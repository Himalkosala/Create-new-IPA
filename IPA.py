import xlwings as xw
#install xlwings before run the script (PS:pip install xlwings)
wb1 = xw.Book('Actual Vs. Budget - LBY contract  Summ IPA No. 08.xlsx') #Give the excel file location, if file is one another dirctory C:\\Users\\..\\.xlsx
wb2 = xw.Book('Actual Vs. Budget - LBY contract  Summ IPA No. 08.xlsx')
ws1 = wb1.sheets(1) #if more than one sheet there list all as wsn=wb1.sheet(n) {n= give each separate number}
ws2 = wb2.sheets(1)


my_values= ws1.range('K11:K521').options(ndim=2).value #Give the range that need to copy
ws2.range('G11:G521').value= my_values #Give the range that need to paste
ws2.range('I11:I521').clear_contents() #Give the range that need to clear
wb2.save('excel.xlsx') #Give the file name to save
wb2.app.quit()
