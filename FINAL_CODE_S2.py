import pandas as pd
import numpy as np
import os

# 01. 엑셀 데이터 불러오기 ############################################################################################
# 파일이 저장된 디렉토리 지정 (예: FINAL_code_s1에서 파일을 저장한 디렉토리)
directory_path =  r"Z:\남수경\quant\watchdog_test"

# 디렉토리 내의 모든 파일 목록 가져오기
files = os.listdir(directory_path)

# 날짜를 포함한 파일 목록 필터링 (QUANT_LARGE_OUTPUT_YYYYMMDD.xlsx 형식의 파일들만)
date_files = [f for f in files if f.startswith('QUANT_LARGE_OUTPUT') and f.endswith('.xlsx')]

# 파일이 없으면 종료
if not date_files:
    print("대상 파일이 디렉토리에 없습니다.")
else:
    # 파일 이름에서 날짜를 추출하여 최신 파일 찾기
    date_strs = [f.split('_')[2].split('.')[0] for f in date_files]  # 날짜 부분만 추출
    latest_date_str = max(date_strs)  # 가장 최신 날짜 선택
    
    # 최신 날짜로 파일 경로 생성
    output_large_file_name = f"QUANT_LARGE_OUTPUT_{latest_date_str}.xlsx"
    output_small_file_name = f"QUANT_SMALL_OUTPUT_{latest_date_str}.xlsx"
    
    # 'FINAL_code_s1'에서 저장된 파일 불러오기
    df_large = pd.read_excel(os.path.join(directory_path, output_large_file_name))
    df_small = pd.read_excel(os.path.join(directory_path, output_small_file_name))

    # 이제 df_large와 df_small을 이용하여 후속 작업을 진행
    print(f"파일을 성공적으로 불러왔습니다: {output_large_file_name}, {output_small_file_name}")

# 02. 필요한 변수 생성 ################################################################################################
def create_variables(df):
    df['T_PER'] = df['종가'] / df['EPS(지배)']
    df['T_PER_C'] = pd.to_numeric(np.where(df['T_PER'] < 0, 0, df['T_PER']), errors='coerce')
    df['F_PER'] = df['종가'] / df['EPS(Fwd.12M, 지배)']
    df['F_PER_C'] = pd.to_numeric(np.where(df['F_PER'] < 0, 0, df['F_PER']), errors='coerce')
    df['순부채_R'] = df['순부채'] * 1000  # 순부채단위 천원 → 원 변경
    df['EV'] = df['시가총액'] + df['순부채_R']  # 기업가치계산
    df['T_EBITDA'] = df['EBITDA'] * 1000  # EBITDA단위 천원 → 원 변경
    df['T_EVEBITDA'] = df['EV'] / df['T_EBITDA']
    df['F_EBITDA'] = df['EBITDA(Fwd.12M)'] * 1000  # EBITDA단위 천원 → 원 변경
    df['F_EVEBITDA'] = df['EV'] / df['F_EBITDA']
    df['T_PBR'] = df['종가'] / df['BPS(지배)']
    df['F_PBR'] = df['종가'] / df['BPS(Fwd.12M, 지배)']
    df['T_PCF'] = df['종가'] / df['FCFPS(Adj.,Wgt.)']
    df['T_PCF_C'] = np.where(df['T_PCF'] > 0, df['T_PCF'], 0)
    df['T_PCF_C'] = pd.to_numeric(df['T_PCF_C'], errors='coerce')
    df['T_SPSG'] = pd.to_numeric(df['SPS증가율(YoY)'], errors='coerce')
    df['F_SPSG'] = pd.to_numeric(df['SPS Growth 1/0 -Y'], errors='coerce')
    df['F_EPS_M'] = pd.to_numeric(df['EPS(Fwd.12M) 변화율(3개월, 지배)'], errors='coerce')
    df['PRICE_M'] = pd.to_numeric(df['3개월전대비수익률'], errors='coerce')
    return df

# LARGE와 SMALL 데이터셋에 각각 변수 생성
large_df = create_variables(df_large)
small_df = create_variables(df_small)

# 03. 사분위수 생성 ###################################################################################################
def create_quantiles(df, variables_above_zero, variables_all):
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

    return df

# LARGE와 SMALL 데이터셋에 각각 사분위수 생성
variables_above_zero = ['T_PER_C', 'F_PER_C', 'T_EVEBITDA', 'F_EVEBITDA', 'T_PBR', 'F_PBR', 'T_PCF_C']
variables_all = ['ATT_PBR', 'ATT_EVIC', 'ATT_PER', 'ATT_EVEBIT', 'T_SPSG', 'F_SPSG', 'F_EPS_M', 'PRICE_M']

large_df = create_quantiles(large_df, variables_above_zero, variables_all)
small_df = create_quantiles(small_df, variables_above_zero, variables_all)

# 04. SCORING ########################################################################################################
def scoring(df, variable_rule1, variable_rule2, variable_rule3):
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
    return df

# LARGE와 SMALL 데이터셋에 각각 scoring 적용
variable_rule1 = ['T_PER_C', 'F_PER_C', 'T_EVEBITDA', 'F_EVEBITDA', 'T_PBR', 'F_PBR', 'T_PCF_C']
variable_rule2 = ['ATT_PBR', 'ATT_EVIC', 'ATT_PER', 'ATT_EVEBIT', 'T_SPSG', 'F_SPSG', 'F_EPS_M', 'PRICE_M']
variable_rule3 = ['PRICE_M']

large_df = scoring(large_df, variable_rule1, variable_rule2, variable_rule3)
small_df = scoring(small_df, variable_rule1, variable_rule2, variable_rule3)

# 05. FINAL ##########################################################################################################
def final_output(df, directory_path, latest_date_str):
    df['시가총액(백만원)'] = np.floor(df['시가총액'] / 1000000)
    selected_df1 = df[['Code', 'Name', '시장구분', 'WICS업종명(대)', 'WICS업종명(중)', 'WICS업종명(소)', '산업업종명', '시가총액(백만원)']]
    selected_df2 = df.loc[:, ['Name'] + list(df.loc[:, 'T_PER_C_score':'RANKING'].columns)]
    combined_data = pd.merge(selected_df1, selected_df2, on='Name')
    
    # 파일 이름에 최신 날짜 추가
    file_name = f"{directory_path}/LARGE_CAP_FINAL_{latest_date_str}.xlsx"
    combined_data.to_excel(file_name, index=False)

# LARGE와 SMALL 데이터셋에 대해 각각 최종 결과 저장
final_output(large_df, directory_path, latest_date_str)
final_output(small_df, directory_path, latest_date_str)
