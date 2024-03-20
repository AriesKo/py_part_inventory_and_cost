import gspread
import re
    
    
if __name__ == '__main__':
    gc = gspread.service_account(filename='csv-read-7fb03efe621f.json')
    # sh = gc.open('BOM').worksheet('FPA-SMB')
    sh = gc.open('BOM').get_worksheet(0)
    print(sh.get('B4'))
    
    # values_list = sh.col_values(1)
    values_list = sh.row_values(1)
    print(values_list)
    
    list_of_lists = sh.get_all_values()
    
    cell = sh.find('J1,J2')
    cell = sh.findall('J1,J2')
    
    amount_re = re.compile('C0603CJ?')
    cell = sh.findall(amount_re)

    
    