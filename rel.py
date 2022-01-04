# from os import startfile
from openpyxl import load_workbook
wb = load_workbook('RELIANCE.xlsx')
sheet = wb.active
sheet=wb['Sheet1']
for row in range(2,10):
    count=0
    openSum=0
    highSum=0
    lowSum=0
    closeSum=0
    adjCloseSum=0
    volumeSum=0
    for i in range(row-10,row+1):
        if i<2:
            continue
        else:
            count+=1
            openSum+=sheet['B'+str(i)].value
            highSum+=sheet['C'+str(i)].value
            lowSum+=sheet['D'+str(i)].value
            closeSum+=sheet['E'+str(i)].value
            adjCloseSum+=sheet['F'+str(i)].value
            volumeSum+=sheet['G'+str(i)].value
    if count>0:
        sheet['H'+str(row)].value=openSum/count
        sheet['I'+str(row)].value=highSum/count
        sheet['J'+str(row)].value=lowSum/count
        sheet['K'+str(row)].value=closeSum/count
        sheet['L'+str(row)].value=adjCloseSum/count
        sheet['M'+str(row)].value=volumeSum/count

wb.save("RELIANCE.xlsx")
# startfile("RELIANCE.xlsx")



