     ### Objectives ###
# 1. Query Karens PB Loan File
# 2. Forward Fill LC LO
# 3. Dropna on ['PB Loan #']
# 4. Query Current PB Loan File
# 5. Format BOTH files' ['PB Loan #'] for effective comparison
# 6. Filter Karen's file for only new PB Loan #s
# 7. Save file to f'./__DATA__/new_pb_loans/{CURRENT_DATE}__pb_loans.xlsx'
# 8. Manually add those new PB Loans to the usable file (manual because naming conventions aren't being followed)

## DATA FRAMES ##
# kf = Karen Frame (Karens PB Loans)
# cf = Current Frame (Current PB Loans)
# nf = New Frame (New PB Loans)

import pandas as pd

CURRENT_DATE = str(pd.to_datetime('today')).split(' ')[0]

kf = pd.read_excel('/Users/jakobbellamy/Dropbox/Regional Accounts/Concession tracking/Copy of Payback Loan Tracking .xlsx', skiprows=1)

kf.drop(columns=['LC LO', 'Unnamed: 15', 'Unnamed: 16', 'Unnamed: 31'], inplace=True)
kf.rename(columns={'Unnamed: 3': 'LC LO'}, inplace=True)
kf['LC LO'] = kf['LC LO'].ffill(axis=0)

kf['PB Loan #'] = kf['PB Loan #'].apply(lambda x: pd.to_numeric(x, errors='coerce'))
kf = kf.dropna(subset=['PB Loan #'])
kf['PB Loan #'] = kf['PB Loan #'].apply(lambda x: str(int(x)))
kf.dropna(how='all', axis=1, inplace=True)

cf = pd.read_excel('./__DATA__/Payback Loans.xlsx')
cf['PB Loan #'] = cf['PB Loan #'].apply(lambda x: pd.to_numeric(x, errors='coerce'))
cf = cf.dropna(subset=['PB Loan #'])
cf['PB Loan #'] = cf['PB Loan #'].apply(lambda x: str(int(x)))


nf = kf[~kf['PB Loan #'].isin(cf['PB Loan #'])]
nf = nf[['Unnamed: 0', 'LC LO', 'LC Loan #'] + list(nf.columns)[3:]]
nf.to_excel(f'./__DATA__/new_pb_loans/{CURRENT_DATE}__pb_loans.xlsx', index=False)
