import csv
import xlrd
import xlwt
from xlutils.copy import copy

src_csv_file_path="Feed2.csv"#csv file path
dst_excel_file_path="products.xlsx"#excel file path

excel_file = xlrd.open_workbook(dst_excel_file_path)

wb = copy(excel_file) 
sheet = wb.get_sheet(0) 


csv_reader = csv.reader(open(src_csv_file_path))

excel_row=[]#each row in excel file
table = excel_file.sheet_by_index(0)#derive first table by index
row_values = table.row_values(1)

# transform 1,0 to true or false
for idx,value in enumerate(row_values):
    if type(value) is int:
        if value:
            row_values[idx]="TRUE"
        else:
            row_values[idx]="FALSE"


#read each row, except 0th row
row_index=-1
for row in csv_reader:
    #if 'LASER' not in row[4] and 'INK' not in row[4]:
    #    continue
    #row_index+=1
    #if row_index==0: 
    #    continue
  
    #print (row)
    excel_row=row_values#derive data of the second row (title in the first row)
    excel_row[3]=row[1]#feed: name -> pro: name
    excel_row[4]=row[7]+","+row[8]#  feed: 'Printer compatitibility', 'Page'  -> pro:ShortDescription
    excel_row[16]=row[6]#feed:'Manufacturer part number' -> pro:'Manufacturerpartnumber'
    excel_row[15]=row[0]#feed:'Dynamic Supplier Code' -> pro: 'SKU'
    excel_row[12]=row[6]#feed:'Manufacturer part number' -> pro:'SeName'
    excel_row[90]=row[9]+";"#feed: Tag -> Pro: ProductTags+";"
    excel_row[89]=row[3]+";"#feed: Manufacturer -> Pro:Manufacturers+";"
    excel_row[88]=row[4] #feed:Category -> Pro:Categories 
    excel_row[91]="C:\Users\Cathy\Desktop\Newcsv\MSY inkstore datafeed\Dynamic_Supplies_Product_Images" + row[4] + ".jpg"
    #if 'LASER' in row[4]:
    #    excel_row[91]="C:\Users\Cathy\Downloads\msy commerce\brother-compatible-dr-2125-drum-unit-up-to-12000-pages.jpg"
    #elif 'INK' in row[4]:
    #    excel_row[91]="C:\Users\Cathy\Downloads\msy commerce\canon-compatible-cli526-photo-black-ink.jpg"
    price=row[2]#
    price=price.replace(",","")
    excel_row[70]=float(price)*1.1*2#price
    

    excel_row[84]=row[5]#weight
    excel_row[85]=row[12]#length
    excel_row[86]=row[13]#width
    excel_row[87]=row[14]#height
    #print (excel_row)  
    for i in range(len( excel_row)):
        #for row_index in range(-1, 800):
        sheet.write(row_index+1,i,excel_row[i])
    print('done row:'+str(row_index))


wb.save('my_workbook.xls')



        
