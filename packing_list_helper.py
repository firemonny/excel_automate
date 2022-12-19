from openpyxl.styles import Border, Side
from openpyxl.utils.cell import column_index_from_string
sheet_detail_index = {'start_offest': 'A',
'start_offest_dash':'B',
'end_offset': 'C',
'item_description': 'D',
'size': 'E',
'Color': 'F',
'item#':'G',
'PC': 'I',
'Q_TY': 'J',
'CTNS':'K',
'配件重/单':'M',
'内盒':'N',
'外箱':'O',
'pvc':'P',
'M.M.W.': 'R',
'N.W': 'T',
'G.W':'W',
'CBM':'Z',
'N.W./pcs':'AB'
}

def detail_index(name, row_num):
    return sheet_detail_index[name]+str(row_num)


def set_border(ws, row):
    cell_range = f"{sheet_detail_index['start_offest']}{row}:{sheet_detail_index['N.W./pcs']}{row}"
    none = Side(border_style="none", color="000000")
    double = Side(border_style="double", color="000000")
    for row in ws[cell_range]:
        for cell in row:
            if cell.column == column_index_from_string(sheet_detail_index['start_offest']):
                cell.border = Border(top=double, left=double, bottom=double)
            elif cell.row == column_index_from_string(sheet_detail_index['N.W./pcs']):
                cell.border = Border(top=double, right=double, bottom=double)
            else:
                cell.border = Border(top=double, bottom=double)
            

def construct_row_detail(worksheet, row, detail, num_offset):
    row_num = str(row)
    worksheet[detail_index('start_offest',row_num)]= num_offset
    worksheet[detail_index('start_offest_dash',row_num)]= " - " 
    item_count = detail['數量']//detail['外箱 PC-IN']
    worksheet[detail_index('end_offset',row_num)] = num_offset+item_count - 1
    worksheet[detail_index('item_description',row_num)] = detail['品名']
    worksheet[detail_index('size',row_num)] = detail['尺寸']
    worksheet[detail_index('Color',row_num)] = detail['客戶產品顏色']
    worksheet[detail_index('item#',row_num)] = detail['ITEM#']
    worksheet[detail_index('Q_TY',row_num)] = detail['數量']
    worksheet[detail_index('配件重/单',row_num)] = detail['配件重/单']
    worksheet[detail_index('内盒',row_num)] = detail['内盒']
    worksheet[detail_index('外箱',row_num)] = detail['外箱']
    worksheet[detail_index('pvc',row_num)] = detail['pvc/单']
    worksheet[detail_index('N.W',row_num)] = detail['N.W.(KG)']
    worksheet[detail_index('G.W',row_num)] = detail['G.W.(KG)']
    worksheet[detail_index('CBM',row_num)] = detail['CBM(M3)']
# 'M.M.W.': 'R'?

def construct_customer_po_info(worksheet,po_info):
    worksheet['D3'] = po_info['客戶別']
    worksheet['D4'] = po_info['採購單號']
    worksheet['T3'] = po_info['採購單日期']
    worksheet['T4'] = f"{po_info['出貨']} {po_info['出貨地']}"

def construct_po_summery(worksheet, row):
    start_row_num = str(6)
    end_row_num = str(row-1)
    row_num = str(row)
    worksheet[detail_index('start_offest',row_num)] = "TOTAL:"
    worksheet[detail_index('Q_TY',row_num)]= f"=SUM({sheet_detail_index['Q_TY']}{start_row_num}:{sheet_detail_index['Q_TY']}{end_row_num})" 
    set_border(worksheet,row)



