# 01. 엑셀 데이터 불러오기 ############################################################################################
import sys
import pandas as pd
import numpy as np
import os
from sklearn.linear_model import LinearRegression

# 엑셀 파일 경로를 명령줄 인자로 받기
file_path = "Z:\남수경\quant\watchdog_test"

df = pd.read_excel(file_path, sheet_name='1차', skiprows=12)

# 02. 데이터 스크리닝 #################################################################################################
## 1) 시가총액 3천억 미만 삭제
df = df[df['시가총액'] >= 300000000000]
## 2) 관리종목, 투자유의, 저유동성, 순부채비율 > 200% 삭제
df['관리종목'] = (df['관리종목여부(1:관리, 0:정상)']+df['거래정지여부(1:정지, 0:정상)']+df['정리매매구분(1:해당, 0:정상)']+df['불성실공시법인구분(1:해당, 0:미해당)'] >= 1).astype(int) #astype(int): 조건문이 TRUE이면 1, FALSE이면 0으로 변환
df['투자유의'] = (df['투자유의구분(1:유의, 0:정상)']+df['투자주의환기종목(코)구분(1:해당, 0:미해당)'] >= 1).astype(int)
df['저유동성'] = (df['저유동성종목구분(1:해당, 0:미해당)'] >= 1).astype(int)
df['순부채비율_C'] = (df['순부채비율'] > 200).astype(int)
df = df[(df['관리종목']==0) & (df['투자유의']==0) & (df['저유동성']==0) & (df['순부채비율_C']==0)]
## 3) 재무정보 누락 종목 삭제 (매출액. 매출총이익, 영업이익)
df = df.dropna(subset=['매출액', '매출총이익', '영업이익'])

# 03. 시총 나누기 #####################################################################################################
df['MKTCAP_C'] = np.where(df['시가총액'].isna(), '',
                 np.where(df['시가총액'] >= 2000000000000, 'Large',
                 np.where(df['시가총액'] >= 300000000000, 'Medium',
                 np.where(df['시가총액'] >= 100000000000, 'Small', 'Others'))))
df = df.sort_values(by='시가총액', ascending=False)

# 04. 데이터 저장 #####################################################################################################
df_large = df[df['MKTCAP_C']=='Large']
df_Small = df[df['MKTCAP_C']=='Medium']

df.to_excel(file_path+'QUANT_S1.xlsx', index=False)
df_large.to_excel(file_path+'QUANT_Large.xlsx', index=False)
df_Small.to_excel(file_path+'QUANT_Small.xlsx', index=False)

######################################################################################################################
### LARGE CAP & SMALL CAP 실행 시 변경 ###
######################################################################################################################
#df = df_large 
df = df_Small 
######################################################################################################################
######################################################################################################################

# 05. PBR&ROE 관련 변수 생성 ##########################################################################################
## ROE_C
df['ROE'] = df['ROE(지배)']
df['ROE_C'] = np.where((df['ROE'] < 0) | (df['ROE'] >= 100), '', df['ROE'])
## ROE_B
df['ROE_B'] = np.where(pd.isna(df['ROE']), 'Y', 'N')
## ROE_R
df['ROE_C'] = pd.to_numeric(df['ROE_C'], errors='coerce')
df['ROE_R'] = df['ROE_C']/100
## PBR_C
df['PBR'] = df['종가']/df['BPS(지배)']
df['PBR_C'] = np.where(df['PBR'] < 0, 0, np.where(df['PBR'] >= 10, np.nan, df['PBR'])).round(6)
## PBR_C_E
x = df[['ROE_R']].values
y = df['PBR_C'].values
valid = ~np.isnan(x).any(axis=1) & ~np.isnan(y)
x_no_nan = x[valid]
y_no_nan = y[valid]
model = LinearRegression()
model.fit(x_no_nan, y_no_nan)
intercept_value = model.intercept_
df['PBR_C_E'] = np.nan

# 조건 1: x와 y 모두 결측값이 없는 경우, 회귀직선으로 예측한 값을 할당
valid_both = df['ROE_R'].notna() & df['PBR_C'].notna()
df.loc[valid_both, 'PBR_C_E'] = model.predict(df.loc[valid_both, ['ROE_R']])
# 조건 2: x 값만 있는 경우, 예측된 회귀직선에 x 값을 넣어서 구한 값을 할당
valid_x_only = df['ROE_R'].notna() & df['PBR_C'].isna()
df.loc[valid_x_only, 'PBR_C_E'] = model.predict(df.loc[valid_x_only, ['ROE_R']])
# 조건 3: x 값이 빈칸이고, 원데이터가 결측값(ROE_B='Y')인 경우, 빈칸을 할당
df.loc[(df['ROE_R'].isna()) & (df['ROE_B']=='Y'), 'PBR_C_E'] = np.nan
# 조건 4: x 값이 빈칸이고, 원데이터가 결측값(ROE_B='N')이 아닌 경우, 절편값으로 할당
df.loc[(df['ROE_R'].isna()) & (df['ROE_B']=='N'), 'PBR_C_E'] = intercept_value
df['PBR_C_E'] = df['PBR_C_E'].round(6)

intercept_value = model.intercept_
slope_value = model.coef_[0]
print(f"회귀식: y = {slope_value:.4f} * x + {intercept_value:.4f}")

## ATT_PBR
df['PBR_C_E'] = pd.to_numeric(df['PBR_C_E'], errors='coerce')
df['PBR_C'] = pd.to_numeric(df['PBR_C'], errors='coerce')
df['ATT_PBR'] = (df['PBR_C_E']/df['PBR_C']-1).round(2)

# 06. EVIC&ROIC 관련 변수 생성 ########################################################################################
## ROIC_C
df['ROIC_C'] = np.where((df['ROIC'] < 0) | (df['ROIC'] >= 400), '', df['ROIC'])
df['ROIC_R'] = pd.to_numeric(df['ROIC_C'], errors='coerce') #숫자형태로 변경
print(df[['ROIC','ROIC_R']]) # 결과 확인
df['ROIC_R2'] = df['ROIC_R']/100 
## ROIC_B
df['ROIC_B'] = np.where(pd.isna(df['ROIC']), 'Y', 'N')
## EVIC_C
pd.options.display.float_format = '{:.0f}'.format
df['순부채_R'] = df['순부채']*1000
df['IC_R'] = df['IC']*1000
df['EV'] = df['시가총액']+df['순부채_R'] #기업가치계산
df['EVIC'] = (df['EV']/df['IC_R']) #EV/IC 계산
df['EVIC_C'] = np.where((df['EVIC'] < 0) | (df['EVIC'] >= 40), '', df['EVIC'])
df['EVIC_C'] = pd.to_numeric(df['EVIC_C'], errors='coerce')
## EVIC_C_E
x = df[['ROIC_R2']].values
y = df['EVIC_C'].values
valid = ~np.isnan(x).any(axis=1) & ~np.isnan(y)
x_no_nan = x[valid]
y_no_nan = y[valid]
model = LinearRegression()
model.fit(x_no_nan, y_no_nan)
intercept_value = model.intercept_
df['EVIC_C_E'] = np.nan

# 조건 1: x와 y 모두 결측값이 없는 경우, 회귀직선으로 예측한 값을 할당
valid_both = df['ROIC_R2'].notna() & df['EVIC_C'].notna()
df.loc[valid_both, 'EVIC_C_E'] = model.predict(df.loc[valid_both, ['ROIC_R2']]) #조건을 만족하는 행들에 대해서만 예측함(model.predict(x))
# 조건 2: x 값만 있는 경우, 예측된 회귀직선에 x 값을 넣어서 구한 값을 할당
valid_x_only = df['ROIC_R2'].notna() & df['EVIC_C'].isna()
df.loc[valid_x_only, 'EVIC_C_E'] = model.predict(df.loc[valid_x_only, ['ROIC_R2']])
# 조건 3: x 값이 빈칸이고, 원데이터가 결측값(ROIC_B='Y')인 경우, 빈칸을 할당
df.loc[(df['ROIC_R2'].isna()) & (df['ROIC_B']=='Y'), 'EVIC_C_E'] = np.nan
# 조건 4: x 값이 빈칸이고, 원데이터가 결측값(ROIC_B='N')이 아닌 경우, 절편값으로 할당
df.loc[(df['ROIC_R2'].isna()) & (df['ROIC_B']=='N'), 'EVIC_C_E'] = intercept_value

intercept_value = model.intercept_
slope_value = model.coef_[0]
print(f"회귀식: y = {slope_value:.4f} * x + {intercept_value:.4f}")

## EV_NDT
df['EV_NDT'] = df['EVIC_C_E']*df['IC_R']-df['순부채_R']
## ATT_EVIC
df['ATT_EVIC'] = (df['EV_NDT']/df['시가총액']-1).round(2)

# 07. PER&EPSG 관련 변수 생성 #########################################################################################
## EPSG_C
df['EPS Growth Fwd.12M/LTM(지배)'] = pd.to_numeric(df['EPS Growth Fwd.12M/LTM(지배)'], errors='coerce')
df['EPSG_C'] = np.where((df['EPS Growth Fwd.12M/LTM(지배)'] < 0) | (df['EPS Growth Fwd.12M/LTM(지배)'] >= 500), '', df['EPS Growth Fwd.12M/LTM(지배)'])
df['EPSG_R'] = pd.to_numeric(df['EPSG_C'], errors='coerce')
df['EPSG_R'] = df['EPSG_R']/100
## EPSG_B
df['EPSG_B'] = np.where(pd.isna(df['EPS Growth Fwd.12M/LTM(지배)']), 'Y', 'N')
## PER_C
pd.options.display.float_format = '{:.0f}'.format
df['PER'] = df['종가']/df['EPS(Fwd.12M, 지배)']
df['PER_C'] = np.where(df['PER'] < 0, 0, np.where(df['PER'] >= 50, np.nan, df['PER'])).round(5) #숫자형이기 때문에 조건에 ''(문자형) 적용 시 오류 발생
## PER_C_E
x = df[['EPSG_R']].values
y = df['PER_C'].values
valid = ~np.isnan(x).any(axis=1) & ~np.isnan(y)
x_no_nan = x[valid]
y_no_nan = y[valid]
model = LinearRegression()
model.fit(x_no_nan, y_no_nan)
intercept_value = model.intercept_
df['PER_C_E'] = np.nan

# 조건 1: x와 y 모두 결측값이 없는 경우, 회귀직선으로 예측한 값을 할당
valid_both = df['EPSG_R'].notna() & df['PER_C'].notna()
df.loc[valid_both, 'PER_C_E'] = model.predict(df.loc[valid_both, ['EPSG_R']]) #조건을 만족하는 행들에 대해서만 예측함(model.predict(x))
# 조건 2: x 값만 있는 경우, 예측된 회귀직선에 x 값을 넣어서 구한 값을 할당
valid_x_only = df['EPSG_R'].notna() & df['PER_C'].isna()
df.loc[valid_x_only, 'PER_C_E'] = model.predict(df.loc[valid_x_only, ['EPSG_R']])
# 조건 3: x 값이 빈칸이고, 원데이터가 결측값(EPSG_B='Y')인 경우, 빈칸을 할당
df.loc[(df['EPSG_R'].isna()) & (df['EPSG_B']=='Y'), 'PER_C_E'] = np.nan
# 조건 4: x 값이 빈칸이고, 원데이터가 결측값(EPSG_B='N')이 아닌 경우, 절편값으로 할당
df.loc[(df['EPSG_R'].isna()) & (df['EPSG_B']=='N'), 'PER_C_E'] = intercept_value

intercept_value = model.intercept_
slope_value = model.coef_[0]
print(f"회귀식: y = {slope_value:.4f} * x + {intercept_value:.4f}")

## ATT_PER 
df['ATT_PER'] = df['PER_C_E']/df['PER_C']-1
# inf 값과 -inf값을 NaN으로 대체
df['ATT_PER'].replace([np.inf, -np.inf], np.nan, inplace=True)

# 08. EVEBIT&EBITG 관련 변수 생성 #####################################################################################
## EBITG_C 
df['EBITG'] = df['EBIT(Fwd.12M)']/df['EBIT']*100-100 #두 값 중 하나만 있어도 -100으로 계산됨
df['EBITG_C'] = np.where((df['EBITG'] < 0) | (df['EBITG'] >= 500), np.nan, df['EBITG'])
df['EBITG_R'] = df['EBITG_C']/100
## EBITG_B
df['EBITG_B'] = np.where(df['EBIT(Fwd.12M)'].isna() & df['EBIT'].isna(), 'Y', 'N') #SO, 두 값 모두 없어야 BLANK가 됨#
## EVEBIT_C
pd.options.display.float_format = '{:.0f}'.format
df['순부채_R'] = df['순부채']*1000
df['EV'] = df['시가총액']+df['순부채_R']
df['EBIT_R'] = df['EBIT(Fwd.12M)']*1000 #EBIT단위 천원 → 원 변경
df['EVEBIT'] = df['EV']/df['EBIT_R']
pd.options.display.float_format = '{:.2f}'.format
## EVEBIT_B
df['EVEBIT_B'] = np.where(df['EVEBIT'].isna(),'Y','N')
## EVEBIT_C
df['EVEBIT_C'] = np.where((df['EVEBIT'] < 0) | (df['EVEBIT'] >= 50), '', df['EVEBIT'])
df['EVEBIT_R'] = pd.to_numeric(df['EVEBIT_C'], errors='coerce') #숫자형태로 변경
## EVEBIT_C_E
x = df[['EBITG_R']].values
y = df['EVEBIT_R'].values
valid = ~np.isnan(x).any(axis=1) & ~np.isnan(y)
x_no_nan = x[valid]
y_no_nan = y[valid]
model = LinearRegression()
model.fit(x_no_nan, y_no_nan)
intercept_value = model.intercept_
df['EVEBIT_C_E'] = np.nan

# 조건 1: x와 y 모두 결측값이 없는 경우, 회귀직선으로 예측한 값을 할당
valid_both = df['EBITG_R'].notna() & df['EVEBIT_R'].notna()
df.loc[valid_both, 'EVEBIT_C_E'] = model.predict(df.loc[valid_both, ['EBITG_R']]) #조건을 만족하는 행들에 대해서만 예측함(model.predict(x))
# 조건 2: x 값만 있는 경우, 예측된 회귀직선에 x 값을 넣어서 구한 값을 할당
valid_x_only = df['EBITG_R'].notna() & df['EVEBIT_R'].isna()
df.loc[valid_x_only, 'EVEBIT_C_E'] = model.predict(df.loc[valid_x_only, ['EBITG_R']])
# 조건 3: x 값이 빈칸이고, 원데이터가 결측값(EBITG_B='Y')인 경우, 빈칸을 할당
df.loc[(df['EBITG_R'].isna()) & (df['EBITG_B']=='Y'), 'EVEBIT_C_E'] = np.nan
# 조건 4: x 값이 빈칸이고, 원데이터가 결측값(EBITG_B='N')이 아닌 경우, 절편값으로 할당
df.loc[(df['EBITG_R'].isna()) & (df['EBITG_B']=='N'), 'EVEBIT_C_E'] = intercept_value

intercept_value = model.intercept_
slope_value = model.coef_[0]
print(f"회귀식: y = {slope_value:.4f} * x + {intercept_value:.4f}")

## EV_NDT2
df['EV_APPLY'] = df['EVEBIT_C_E'].fillna(0).replace([np.inf, -np.inf], 0) * df['EBIT_R'].fillna(0).replace([np.inf, -np.inf], 0)
df['EV_APPLY'] = df['EV_APPLY'].astype(int)
df['EV_NDT2'] = df['EVEBIT_C_E']*df['EBIT_R']-df['순부채_R'].astype(int)
## ATT_EVIC
df['ATT_EVEBIT'] = df['EV_NDT2']/df['시가총액']-1

# 09. DATA 출력
# 파일 이름에서 날짜 부분만 추출 (예: 20240912)
file_name = os.path.basename(file_path)  # 파일명만 추출
date_str = file_name.split('__')[0]  # '__' 앞부분을 날짜로 추출

# 새로운 파일 이름 만들기 (예: 'QUANT_SMALL_OUTPUT_20240912.xlsx')
output_large_file_name  = f"QUANT_LARGE_OUTPUT_{date_str}.xlsx"
output_small_file_name  = f"QUANT_SMALL_OUTPUT_{date_str}.xlsx"

# 04. 결과를 엑셀 파일로 저장 (Large 파일과 Small 파일을 각각 저장)
df.to_excel(output_large_file_name, index=False)
df.to_excel(output_small_file_name, index=False)

print(f"결과 파일이 저장되었습니다: {output_large_file_name} 및 {output_small_file_name}")