import xlsxwriter

#create file (outWorkbook) and worksheet (outSheet)
outWorkbook = xlsxwriter.Workbook('out.xlsx')
outSheet = outWorkbook.add_worksheet()

# declare data
names = ['Alice', 'Bob', 'Charlie']
ages = [30, 25, 35]

# write headers
outSheet.write('A1', 'Name')
outSheet.write('B1', 'Age')

# write data to file
for item in range(len(names)):
    outSheet.write(item+1, 0, names[item])  # Column A
    outSheet.write(item+1, 1, ages[item])   # Column B

outSheet.write('D1', 'Total')
outSheet.write_formula('D2', '=sum(B2:B5)')  # Write total of ages in D2   
outWorkbook.close()  # Close the workbook to save changes
print("Data written to out.xlsx successfully.")