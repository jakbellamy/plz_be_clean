import pandas as pd
import numpy as np

fundings_path = './__DATA__/Funding By Referral Source (2006) (10).xlsx'
branch_details_path = './__DATA__/Branch Production Details (3073) 20211118.xlsx'

def read_fundings(path):
    df = pd.read_excel(path, sheet_name='Details', skiprows=4)
    df.columns = [col.split('\n')[0] for col in df.columns]
    df['Loan Number'] = df['Loan Number'].apply(str)
    df.set_index('Loan Number', inplace=True)
    return df

def read_branch_details(path):
    df = pd.read_excel(path, sheet_name='RAW DATA', skiprows=4)
    df['Loan Number'] = df['Loan Number'].apply(str)
    df.set_index('Loan Number', inplace=True)
    return df

def load_data(fundings_path, branch_details_path):
    fundings = read_fundings(fundings_path)
    branch_details = read_branch_details(branch_details_path)
    branch_details = branch_details[branch_details['Activity Type'] == 'Fundings']
    return branch_details.join(fundings[['Referral Name', 'Referral Source']])

df = load_data(fundings_path, branch_details_path)

def search_address(substr):
    return df[df['Subject Property Address'].apply(lambda x: substr.lower() in x.lower())]

# Now search and add.
selected_loans = []

def add(loan_number):
    selected_loans.append(str(loan_number))

def return_loans():
    return df.loc[selected_loans][[
        'Last Name', 'Funding Date Time', 'Loan Officer', 'Property Type', 'Subject Property Address', 
        'Subject Property City', 'Subject Property State', 'Subject Property Zip', 'Buyers Agent Contact Name', 
        'Buyers Agent Name', 'Referral Source', 'Referral Name']]