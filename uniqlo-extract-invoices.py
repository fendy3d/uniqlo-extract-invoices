import os
import pdfplumber
import pandas as pd
import csv

pathToPdfs = os.getcwd()+"/dropPdfHere/"
header_row = [
'Nomor Invoice',
'Date of Invoice',
'Bill to',
'Description',
'DPP',
'PPN',
'Total',
]
list_of_rows = []

def reformatMonth(month):
	if month == "January":
		month = '01'
	elif month == "February":
		month = '02'
	elif month == "March":
		month = '03'
	elif month == "April":
		month = '04'
	elif month == "May":
		month = '05'
	elif month == "June":
		month = '06'
	elif month == "July":
		month = '07'
	elif month == "August":
		month = '08'
	elif month == "September":
		month = '09'
	elif month == "October":
		month = '10'
	elif month == "November":
		month = '11'
	elif month == "December":
		month = '12'
	return month
		

for _, _, files in os.walk(pathToPdfs):
    for filename in files:
        if '.pdf' in filename:
            print ("Extracting " + filename)
            pdf = pdfplumber.open(pathToPdfs + filename)
            page = pdf.pages[0] # get the page
            
            texts = page.extract_text()
            table = page.extract_tables()[0]
            list_of_texts = texts.split('\n')

            
            if 'AFI-' in filename:
            	print("==============================")
            	print("AFI invoices - " + filename)
            	nomor_invoice = list_of_texts[6].split(': ')[-1]
            	date_of_invoice_list = list_of_texts[5].split(', ')[-1].split(' ')
            	day = str(date_of_invoice_list[0])
            	month = reformatMonth(date_of_invoice_list[1])
            	year = str(date_of_invoice_list[2])
            	date_of_invoice = day + '/' + month + '/' + year
            	bill_to = list_of_texts[10].split('bellow account ')[-1]
            	description = "XXX"
            	dpp = "XXX"
            	ppn = "XXX"
            	total = table[-1][-1].split(' ')[-1].replace('.','')

            elif 'PBI-' in filename:
            	print("==============================")
            	print("PBI invoices - " + filename)
            	
            	nomor_invoice = list_of_texts[2].split('Invoice No : ')[-1]
            	date_of_invoice_list = list_of_texts[1].split('Date : ')[-1].split(' ')
            	day = str(date_of_invoice_list[1].replace(',', ''))
            	month = reformatMonth(date_of_invoice_list[0])
            	year = str(date_of_invoice_list[2])
            	date_of_invoice = day + '/' + month + '/' + year
            	bill_to = list_of_texts[2].split(' Invoice No')[0]
            	description = list_of_texts[7].split(' Rp')[0]
            	dpp = list_of_texts[8].split(' ')[-1].replace('.','')
            	ppn = list_of_texts[9].split(' ')[-1].replace('.','')
            	total = list_of_texts[10].split(' ')[-1].replace('.','')

            list_of_rows.append([
                    nomor_invoice,
                    date_of_invoice,
                    bill_to,
                    description,
                    dpp,
                    ppn,
                    total,
                    ])

            # counter = 0
            # for item in table:
            # 	print(counter)
            # 	print(item)
            # 	counter += 1
            
    #Exporting csv
    df = pd.DataFrame(list_of_rows, columns=header_row)
    df.to_csv('output-3.csv',index=False)

    print("Success! All hail Lord Fendy!")
