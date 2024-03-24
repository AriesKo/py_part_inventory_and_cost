import xlwings as xw
import re
import pandas as pd

if __name__ == '__main__':
    # "11" : all, "01": cost, "10": inventory
    mode_value = "11"    
    bom_path = r'example/bom_fpa.xlsx'
    xw_cost_path = r'example/cost.xlsx'
    xw_inventory_path = r'example/inventory.xlsx'

    # Excel File 읽기
    if (mode_value == "11"):
        xw_bom = xw.Book(bom_path)
        xw_cost = xw.Book(xw_cost_path)
        xw_inventory = xw.Book(xw_inventory_path)
        
        # 특정 sheet 읽기
        sh_bom = xw_bom.sheets(1)
        sh_cost = xw_cost.sheets(1)
        sh_inventory = xw_inventory.sheets(1)
        
        # Dataframe으로 변환
        df_bomlist = sh_bom.range('A1').options(pd.DataFrame, index=False, expand='table').value
        df_cost = sh_cost.range('A1').options(pd.DataFrame, index=False, expand='table').value
        df_inventory = sh_inventory.range('A1').options(pd.DataFrame, index=False, expand='table').value

        #공백제거
        df_bomlist['Part No.'] = df_bomlist['Part No.'].str.replace(" ","")
        df_cost['품명'] = df_cost['품명'].str.replace(" ","")
        df_inventory['자재명'] = df_inventory['자재명'].str.replace(" ","")

    elif (mode_value == "01"):
        xw_bom = xw.Book(bom_path)
        xw_cost = xw.Book(xw_cost_path)
        
        # 특정 sheet 읽기
        sh_bom = xw_bom.sheets(1)
        sh_cost = xw_cost.sheets(1)
        
        # Dataframe으로 변환
        df_bomlist = sh_bom.range('A1').options(pd.DataFrame, index=False, expand='table').value
        df_cost = sh_cost.range('A1').options(pd.DataFrame, index=False, expand='table').value
        
            #공백제거
        df_bomlist['Part No.'] = df_bomlist['Part No.'].str.replace(" ","")
        df_cost['품명'] = df_cost['품명'].str.replace(" ","")
        
    else:
        xw_bom = xw.Book(bom_path)
        xw_inventory = xw.Book(xw_inventory_path)
        
        # 특정 sheet 읽기
        sh_bom = xw_bom.sheets(1)
        sh_inventory = xw_inventory.sheets(1)
    
        # Dataframe으로 변환
        df_bomlist = sh_bom.range('A1').options(pd.DataFrame, index=False, expand='table').value
        df_inventory = sh_inventory.range('A1').options(pd.DataFrame, index=False, expand='table').value
    
        #공백제거
        df_bomlist['Part No.'] = df_bomlist['Part No.'].str.replace(" ","")
        df_inventory['자재명'] = df_inventory['자재명'].str.replace(" ","")
    
    # BOM에서 필요한 Column만 추출 및 생성
    df_partname = df_bomlist["Part No."].dropna()
    df_qty = df_bomlist["Q'TY"].dropna().astype(int)
    df_bom = pd.concat([df_partname,df_qty], ignore_index=True, axis=1)
    df_bom = df_bom.reset_index(drop=True)
    df_bom = df_bom.rename(columns = {0:'품명', 1:'사용수량'})
    df_bom['단가'] = [0 for i in range(len(df_bom))]
    df_bom['현재고'] = [0 for i in range(len(df_bom))]
    df_bom['출고제한수량'] = [0 for i in range(len(df_bom))]
    df_bom['프로젝트번호'] = [0 for i in range(len(df_bom))]

    # 단가 추가
    if (mode_value =="11" or mode_value == "01"):
        for k in range(len(df_bom)):
            datafilter = df_cost['품명'].notnull() & df_cost['품명'].str.contains(str(df_bom.iloc[k,0])[:],case=False) # 대소문자 구분X
            if len(df_cost.loc[datafilter,'최종결산월재고단가'].values) == 0:
                df_bom.loc[k,["단가"]] = 0
            elif len(df_cost.loc[datafilter,'최종결산월재고단가'].values) > 1 : # 동일한 부품단가가 있는 경우 마지막값 사용
                df_bom.loc[k,["단가"]] = int(df_cost.loc[datafilter,'최종결산월재고단가'].values[-1])
            else:
                df_bom.loc[k,["단가"]] = int(df_cost.loc[datafilter,'최종결산월재고단가'].values)
    
    # 재고 추가
    if (mode_value =="11" or mode_value == "10"):
        for k in range(len(df_bom)):
            datafilter = df_inventory['자재명'].notnull() & df_inventory['자재명'].str.contains(str(df_bom.iloc[k,0])[:],case=False) # 대소문자 구분X
            if len(df_inventory[datafilter]['프로젝트번호'].values) == 0:
                pass
            else:
                df_bom.loc[k,['현재고']] = str(df_inventory.loc[datafilter,'현재고'].values)
                df_bom.loc[k,['출고제한수량']] = str(df_inventory.loc[datafilter,'출고제한수량'].values)
                df_bom.loc[k,['프로젝트번호']] = str(df_inventory.loc[datafilter,'프로젝트번호'].values)
            
    print(df_bom)
    
    # CSV 출력
    df_bom.to_csv("result.csv",encoding='cp949')