import pandas as pd 
from openpyxl import load_workbook
from openpyxl.styles import Border
from packing_list_helper import *

# Read the data file
rdc_df = pd.read_csv("data/rdc.csv")
item_list_df = pd.read_csv("data/item_list.csv")
po_details_df = pd.read_csv("data/po_details.csv")

rdc_df.rename(columns = {'出貨地代碼':'出貨'}, inplace = True)

# Modify the po detail csv
po_details_df.rename(columns = {'Item#':'ITEM#'}, inplace = True)
po_details_df.iloc[0]

# handle the item_list
item_list_df["N.W.(KG)"] = item_list_df["G.W.(KG)"] - item_list_df["外箱"]
item_list_df["pvc/单"] = (item_list_df["N.W.(KG)"] / item_list_df["PC-IN"]) - item_list_df["配件重/单"] - item_list_df["内盒"]

# Data join to detail
new_po_detail_df = po_details_df.join(item_list_df.set_index('ITEM#'), on='ITEM#')
new_po_detail_df = new_po_detail_df.join(rdc_df.set_index('出貨'), on='出貨')
new_po_detail_df.columns

# Construct the po_detail report
new_po_detail_df_group = new_po_detail_df.groupby(['客戶PO單號'])
wb = load_workbook('xlsx_template/packing_list_template.xlsx')
temple_sheet = wb.active
for name, group in new_po_detail_df_group:

    target = wb.copy_worksheet(temple_sheet)
    construct_customer_po_info(target,group.iloc[0])
    offset = 1
    row_count = 6
    
    for index, row in group.iterrows():
        construct_row_detail(target, row_count, row, offset)
        offset += row['數量']//row['外箱 PC-IN']
        row_count += 1
    construct_po_summery(target, row_count)
    target.title = str(name)
    
wb.remove_sheet(temple_sheet)
wb.save(f'results/details_po.xlsx')
wb.close