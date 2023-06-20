# -*- coding: utf-8 -*-
"""
Created on Tue Jun 20 09:32:22 2023

@author: bgw60
"""

import pandas as pd


def find_excel_file(file_name, month):
    excel_file = file_name
    sheet_name=month
    
    excel_data = pd.read_excel(excel_file, sheet_name = sheet_name)
    return excel_data



def find_cells_with_word(file_name, sheet_name, search_word):
    # 엑셀 파일 로드
    df = pd.read_excel(file_name, sheet_name=sheet_name, engine='openpyxl',usecols='A:AH')
    
    # 단어가 포함된 셀 찾기
    filtered_df = df[df.apply(lambda row: search_word in ' '.join(map(str, row)), axis=1)]
    
    # 결과 반환
    return filtered_df

def load_date(file_name,sheet_name,keyword):
    excel_file = file_name
    
    excel_data = pd.read_excel(excel_file, sheet_name = sheet_name,header = None)
    
    #print(excel_data)
    month_value = excel_data.at[0,2]
    date_value=[]
    day_value=[]
    work_value=[]
    #일자 데이터 입력
    day_length =2
    row_index = 1
    while True:
        try:
            value = excel_data.iat[row_index, day_length]
            if pd.isnull(value):
                break
            date_value.append(value)
            day_length += 1
        except IndexError:
            break
        
    day_value = excel_data.iloc[2,2:day_length]
    #print(day_value)
    day_value = day_value.tolist()
    #print(len(day_value))
    
    
    for index, row in excel_data.iterrows():
        if row.astype(str).str.contains(keyword, case=False).any():
           
            
            work_value=row.iloc[2:day_length]
            work_value = work_value.fillna('')
            work_value=work_value.tolist()
            
            print(len(work_value))
        elif row.isnull().all():
            break
    
    print(work_value)
    print(day_value)
    print(date_value)
    
    return month_value,date_value, day_value, date_value
    
    
    '''
    #요일 데이터 입력
    column_index =2
    row_index = 2
    while True:
        try:
            value = excel_data.iat[row_index, column_index]
            if pd.isnull(value):
                break
            day_value.append(value)
            column_index += 1
        except IndexError:
            
            break
    column_index =0
    row_index=0
         
    keyword = '김예'
 
    for index, row in excel_data.iterrows():
        if row.astype(str).str.contains(keyword, case=False).any():
           
            
            work_value=row.tolist()
            print(work_value)
        elif row.isnull().all():
            break
    
    nan_indices = [i for i, value in enumerate(work_value) if pd.isna(value)]
   # print(nan_indices)
    if nan_indices:
        sliced_data = work_value[nan_indices[0]+1:nan_indices[1]]
    else:
        ValueError
    work_value = sliced_data
    print(len(work_value))
    print(len(day_value))
    print(len(date_value))
    '''



   
    
        
#result = load_date('수의사근무표_2023(최신2).xlsx', '7월')
#print(result)
     

        
        
