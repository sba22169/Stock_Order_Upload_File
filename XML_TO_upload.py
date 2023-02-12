
import xml.etree.ElementTree as et
import pandas as pd
import numpy as np

def get_po(root):
	""" gets po number from elementtree file"""
	for elem in xroot.iter('ordernumber'):
		po = elem.text
	return po


# function to Save dataframe to excel
def writ_exel(df):
    """ function takes a dataframe as an arg and save it out as an excel file with formatting"""
    writer = pd.ExcelWriter(f"{supplier}_{po}.xlsx", engine='xlsxwriter')
    # Convert the dataframe to an XlsxWriter Excel object. dont want index , so set to false
    df1.to_excel(writer, sheet_name='Order', index=False)
    # Get the xlsxwriter workbook and worksheet objects. Name the worksheet
    workbook  = writer.book
    worksheet = writer.sheets['Order']
    # Add some cell formats. Format the cell to Number, else upload cant read
    format1 = workbook.add_format({'num_format': '0'})
    format2 = workbook.add_format({'num_format': '0'})
    # Note: It isn't possible to format any cells that already have a format such
    # as the index or headers or any cells that contain dates or datetimes.
    # Set the column width and format.
    worksheet.set_column(0, 0, 25, format1)
    # Set the format but not the column width.
    worksheet.set_column(0, 1, None, format2)
    # Close the Pandas Excel writer and output the Excel file.
    writer.close()

def get_supplier(root):
	""" Function takes elementtree root as an arg and returns the supplier name"""
	for item in xroot.findall("./rep/name"):
		if item.text == 'Wella':
			supplier = 'Wella'
		elif item.text == 'Loreal':
			supplier = 'Loreal'
		else:
			supplier = 'Unknown'
		return supplier

def get_stock_qty(root):
	""" Function takes elementtree root as an arg and returns dictionary with stock no and qyt"""
	stock_no = []
	qty = []
	for item in xroot.iter('stock'):
		stock_no.append(item.text)
		for item in xroot.iter('quantity'):
			qty.append(item.text)
	return dict(zip(stock_no, qty))



# The parsing of our “Untitled.xml” file starts at the root of the tree
# namely the <data> element, which contains the entire data structure.
xtree = et.parse("Untitled.xml")
xroot = xtree.getroot()
supplier = get_supplier(xroot)
po = get_po(xroot)

order_dict = get_stock_qty(xroot)

# Make the dataframe and save it to format required by supplier for upload process 

df1 = pd.DataFrame(list(order_dict.items()))
# wella file requires specific headers EAN and quantity
if supplier == "Wella":
    df1.columns = ['EAN', 'quantity']
    df1['EAN'] = df1['EAN'].astype('int64')
    df1['quantity'] = df1['quantity'].astype('int64')
    writ_exel(df1)

elif supplier == "Loreal":
	# Loreal file must be csv with no header or index data
	df1.to_csv(f"{supplier}_{po}.csv",header=False,index=False)



