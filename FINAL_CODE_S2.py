# 01. 엑셀 데이터 불러오기 ############################################################################################
import pandas as pd
import numpy as np

#df = pd.read_excel('C:/Python_coding/FINAL/QUANT_LARGE_OUTPUT_20243Q.xlsx')
df = pd.read_excel('C:/Python_coding/FINAL/QUANT_SMALL_OUTPUT_20243Q.xlsx')

# 02. 필요한 변수 생성 ################################################################################################
df['T_PER'] = df['종가']/df['EPS(지배)']
df['T_PER_C'] = pd.to_numeric(np.where(df['T_PER']<0, 0, df['T_PER']), errors='coerce')
df['F_PER'] = df['종가']/df['EPS(Fwd.12M, 지배)']
df['F_PER_C'] = pd.to_numeric(np.where(df['F_PER']<0, 0, df['F_PER']), errors='coerce')
df['순부채_R'] = df['순부채']*1000 #순부채단위 천원 → 원 변경
df['EV'] = df['시가총액']+df['순부채_R'] #기업가치계산
df['T_EBITDA'] = df['EBITDA']*1000 #EBITDA단위 천원 → 원 변경
df['T_EVEBITDA'] = df['EV']/df['T_EBITDA']
df['F_EBITDA'] = df['EBITDA(Fwd.12M)']*1000 #EBITDA단위 천원 → 원 변경
df['F_EVEBITDA'] = df['EV']/df['F_EBITDA']
df['T_PBR'] = df['종가']/df['BPS(지배)']
df['F_PBR'] = df['종가']/df['BPS(Fwd.12M, 지배)']
df['T_PCF'] = df['종가']/df['FCFPS(Adj.,Wgt.)']
df['T_PCF_C'] = np.where(df['T_PCF']>0, df['T_PCF'], 0)
df['T_PCF_C'] = pd.to_numeric(df['T_PCF_C'], errors='coerce')
df['T_SPSG'] = pd.to_numeric(df['SPS증가율(YoY)'], errors='coerce')
df['F_SPSG'] = pd.to_numeric(df['SPS Growth 1/0 -Y'],errors='coerce')
df['F_EPS_M'] = pd.to_numeric(df['EPS(Fwd.12M) 변화율(3개월, 지배)'], errors='coerce')
df['PRICE_M'] = pd.to_numeric(df['3개월전대비수익률'], errors='coerce')
select_var = df[['Name', 'T_PER_C', 'F_PER_C', 'T_EVEBITDA', 'F_EVEBITDA', 'T_PBR', 'F_PBR', 'T_PCF_C', 'ATT_PBR', 'ATT_EVIC', 'ATT_PER', 'ATT_EVEBIT', 
                 'T_SPSG', 'F_SPSG', 'F_EPS_M', 'PRICE_M']]
select_var.to_excel('OUTPUT_FINAL1.xlsx', index=False)
df['EPSG_R'] = pd.to_numeric(df['EPSG_C'], errors='coerce') 

# 03. 사분위수 생성 ###################################################################################################
variables_above_zero = ['T_PER_C', 'F_PER_C', 'T_EVEBITDA', 'F_EVEBITDA', 'T_PBR', 'F_PBR', 'T_PCF_C'] # 0이상
variables_all = ['ATT_PBR', 'ATT_EVIC', 'ATT_PER', 'ATT_EVEBIT', 'T_SPSG', 'F_SPSG', 'F_EPS_M', 'PRICE_M']

for column in variables_above_zero:
    filtered = df[column][(df[column] > 0) & (df[column].notna())]  # 0 초과 및 BLANK 제외
    df[f'{column}_Q1'] = filtered.quantile(0.25)
    df[f'{column}_Q2'] = filtered.quantile(0.5)
    df[f'{column}_Q3'] = filtered.quantile(0.75)

for column in variables_all:
    filtered = df[column][df[column].notna()]  # BLANK 제외
    df[f'{column}_Q1'] = filtered.quantile(0.25)
    df[f'{column}_Q2'] = filtered.quantile(0.5)
    df[f'{column}_Q3'] = filtered.quantile(0.75)

# 04. SCORING ########################################################################################################
variable_rule1 = ['T_PER_C', 'F_PER_C', 'T_EVEBITDA', 'F_EVEBITDA', 'T_PBR', 'F_PBR', 'T_PCF_C']
variable_rule2 = ['ATT_PBR', 'ATT_EVIC', 'ATT_PER', 'ATT_EVEBIT', 'T_SPSG', 'F_SPSG', 'F_EPS_M', 'PRICE_M']
variable_rule3 = ['PRICE_M']

for column in variable_rule1:
    df[f'{column}_score'] = np.nan

    df[f'{column}_score'] = df.apply(
        lambda row: (
            0 if pd.isna(row[column]) or row[column] <= 0 else 
            1 if row[column] >= row[f'{column}_Q3'] else
            2 if row[f'{column}_Q2'] <= row[column] < row[f'{column}_Q3'] else
            3 if row[f'{column}_Q1'] <= row[column] < row[f'{column}_Q2'] else
            4
        ), axis=1
    )

for column in variable_rule2:
    df[f'{column}_score'] = df.apply(
        lambda row: (
            4 if row[column] >= row[f'{column}_Q3'] else
            3 if row[f'{column}_Q2'] <= row[column] < row[f'{column}_Q3'] else
            2 if row[f'{column}_Q1'] <= row[column] < row[f'{column}_Q2'] else
            1 if row[column] < row[f'{column}_Q1'] else
            0
        ), axis=1
    )

for column in variable_rule3:
    df[f'{column}_score'] = df.apply(
        lambda row:(
            0 if pd.isna(row[column]) else
            1 if row[column] >= row[f'{column}_Q3'] else
            2 if row[f'{column}_Q2'] <= row[column] < row[f'{column}_Q3'] else
            3 if row[f'{column}_Q1'] <= row[column] < row[f'{column}_Q2'] else 
            4
        ), axis=1
    )

df['TOTAL_SCORE'] = (
    (df['T_PER_C_score'] * 0.05) + (df['F_PER_C_score'] * 0.05) +
    (df['T_EVEBITDA_score'] * 0.05) + (df['F_EVEBITDA_score'] * 0.05) +
    (df['T_PBR_score'] * 0.05) + (df['F_PBR_score'] * 0.05) +
    (df['T_PCF_C_score'] * 0.05) + (df['ATT_PBR_score'] * 0.05) +
    (df['ATT_EVIC_score'] * 0.05) + (df['ATT_PER_score'] * 0.1) +
    (df['ATT_EVEBIT_score'] * 0.1) + (df['T_SPSG_score'] * 0.1) +
    (df['F_SPSG_score'] * 0.1) + (df['F_EPS_M_score'] * 0.1) +
    (df['PRICE_M_score'] * 0.05)
)
df['RANKING'] = df['TOTAL_SCORE'].rank(method='dense', ascending=False).astype(int)
df = df.sort_values(by='RANKING')

# 04. FINAL ##########################################################################################################
df['시가총액(백만원)'] = np.floor(df['시가총액'] / 1000000)
selected_df1 = df[['Code', 'Name', '시장구분', 'WICS업종명(대)', 'WICS업종명(중)', 'WICS업종명(소)', '산업업종명', '시가총액(백만원)']]
selected_df2 = df.loc[:, ['Name'] + list(df.loc[:, 'T_PER_C_score':'RANKING'].columns)]
combined_data = pd.merge(selected_df1, selected_df2, on='Name')
#combined_data.to_excel('C:/Python_coding/FINAL/LARGE_CAP_FINAL_20243Q.xlsx', index=False)
combined_data.to_excel('C:/Python_coding/FINAL/SMALL_CAP_FINAL_20243Q.xlsx', index=False)