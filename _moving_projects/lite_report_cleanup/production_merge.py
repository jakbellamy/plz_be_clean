import pandas as pd
import numpy as np

CURRENT_FILE_PATH = './__DATA__/LITE DATA FROM REPORT 2021-09-15.xlsx'
MYREPORTS_FILE_PATH = './__DATA__/Funding By Referral Source (2006) (10).xlsx'

# Load Current Data
current = pd.read_excel(CURRENT_FILE_PATH)
current.drop(columns=['is 2020'], inplace=True)
current.rename(columns={' ': 'Loan Type'}, inplace=True)

# Load MyReports Data
dallas = pd.read_excel(MYREPORTS_FILE_PATH, sheet_name='Details', skiprows=4)
dallas.columns = [col.split('\n')[0] for col in dallas.columns]

# Clean loan numbers for easy merge
dallas['Loan Number'] = dallas['Loan Number'].apply(str)
current['Loan Number'] = current['Loan Number'].apply(str)

# Remove loans that aren't in dallas' list
    # -- Only one loan found: No referral associated
current = current.loc[current['Loan Number'].isin(dallas['Loan Number'])]

# Add loans that are in dallas' list but not in my current report
    # -- most of these should be non-asa since i culled them for early months
dallas_cut = dallas[dallas['Funded Month'] >= pd.to_datetime('2019-01-01')]
missing_loans_from_current = dallas_cut[~dallas_cut['Loan Number'].isin(current['Loan Number'])]
missing_loans_from_current = missing_loans_from_current.drop(columns=[col for col in missing_loans_from_current if not col in current.columns])

# Merge the two datasets   
merged = pd.concat([current.set_index('Loan Number'), missing_loans_from_current.set_index('Loan Number')])

# Pull out null referral sources from merged
merged_na = merged[merged['Referral Source'].isnull()].reset_index()
# Drop real nulls (null in dallas)
dallas_na = dallas[dallas['Referral Source'].isnull()].reset_index()
merged_na = merged_na[~merged_na['Loan Number'].isin(dallas_na['Loan Number'])]
dallas_corrected = dallas[dallas['Loan Number'].isin(merged_na['Loan Number'])]