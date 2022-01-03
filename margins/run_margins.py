import pandas as pd
from ToolSalad.Environment import Environment
from ToolSalad.Database import Database
import glob
import os

data_folder = '/Users/jakobbellamy/Dev/plz_be_clean/margins/__DATA__/'

class MarginsReport(object):
    def __init__(self):
        self.env=Environment()
        self.db=Database(self.env.db_margins)
        self.raw_report_path=max(glob.glob(data_folder + 'raw/*.xlsx'), key=os.path.getctime)
        self.raw_report=pd.read_excel(self.raw_report_path, skiprows=2),
        self.constants_path=data_folder + 'constants/'
        self.payback_loans=pd.read_excel(data_folder + 'Payback Loans.xlsx')

        self.setup()

    pd.set_option('display.max_columns', None)  # change this to a root file that's always imported
    pd.options.mode.chained_assignment = None  # default='warn'

    def setup(self):
        self.load_constants()
        self.load_database()

        self.payback_loans['PB Loan #'] = self.payback_loans['PB Loan #'].apply(
            lambda x: pd.to_numeric(x, errors='coerce'))
        
        if isinstance(self.raw_report, tuple):
            self.raw_report = self.raw_report[0]
        
        df = self.raw_report.dropna(subset=['Funded Date', 'GBR', 'Loan Officer'])
        # df['GBR'] = df['GBR'].apply(lambda x: pd.to_numeric(x, errors='coerce'))
        df = df[df['GBR'] > 0]
        df['Loan Number'] = df['Loan Number'].apply(pd.to_numeric)
        df['Loan Officer'] = df['Loan Officer'].apply(
            lambda x: self.rename_on_key(x, self.name_key, 'Alias'))
        df['Product Code'] = df['Product']
        df['Product'] = df['Product'].apply(
            lambda x: self.rename_on_key(x,
                                        self.loan_types,
                                        'Coding', remove_white_space=True))
        df['Units'] = 1  # For more reliable aggregation than df.shape[0]
        df['GBR - Target Dollars'] = df.apply(
            lambda x: self.calc_revenue_from_margin(
                x['GBR - Target'], x['Amount']),
            axis=1)

        df = df[df['Loan Officer'].apply(lambda x: isinstance(x, str))]

        df['Executing Loan Officer'] = df.apply(self.assign_executing_officer_to_pbl, axis=1)
        df['GBR - Adjusted Dollars'] = df.apply(self.remove_pbl_commissions, axis=1)

        df['GBR - Adjusted Dollars'] = df.apply(self.adjust_team_rev, axis=1)
        df['Executing Loan Officer'] = df['Executing Loan Officer'].apply(
            lambda x: self.team_key[x] if x in [*self.team_key] else x)

        df['Loan Officer'] = df['Executing Loan Officer']
        df.drop(columns=['Executing Loan Officer'], inplace=True)

        df['Net Income'] = df.apply(self.calc_net_income, axis=1)

        self.report = df

    def load_constants(self):
        with open(self.constants_path + 'valid_payback_officers.txt') as file:
            self.valid_pb_officers = list(filter(bool, file.read().split('\n')))

        with open(self.constants_path + 'team_key.txt') as file:
            self.team_key = list(filter(bool, file.read().split('\n')))
            self.team_key = {x.split('=>')[0]: x.split('=>')[1] for x in self.team_key}
    
    def load_database(self):
        self.loan_types = self.db.fetch_table('Loan Types')
        self.name_key = self.db.fetch_table('Name Key')
        self.commissions = self.db.fetch_table('Commission Rates').drop_duplicates()
    
    def assign_executing_officer_to_pbl(self, row):
            if row['Loan Number'] in self.payback_loans['PB Loan #'] \
                    and row['Loan Officer'] in self.valid_pb_officers:
                pb_loan = self.payback_loans.set_index('PB Loan #').loc[row['Loan Number']]
                return pb_loan['LC LO']
            else:
                return row['Loan Officer']

    def remove_pbl_commissions(self, row):
        try:
            if row['Loan Number'] in list(self.payback_loans['PB Loan #']) \
                    and row['Loan Officer'] in self.valid_pb_officers:
                pb_loan = self.payback_loans.set_index('PB Loan #').loc[row['Loan Number']]
                try:
                    com_agreement = self.commissions.set_index('Loan Officer').loc[pb_loan['LC LO']]
                    commission_bps = com_agreement['Commission BPS']
                    max_commission = com_agreement['Maximum']
                except:
                    print('No Commission found for: ', pb_loan['LC LO'])
                    commission_bps = 100  # Default if not found
                    max_commission = 6000  # Default if not found

                lo_commission = row['Amount'] * (commission_bps / 10000)
                if lo_commission <= max_commission:
                    return lo_commission + row['GBR']
                else:

                    return max_commission + row['GBR']
            else:
                return row['GBR']
        except Exception as e:
            print('Failed on \n', row, '\n', e, '\n-------------------')

    def adjust_team_rev(self, row):
        if row['Loan Officer'] in [*self.team_key]:
            team_lead = self.team_key[row['Loan Officer']]
            com_agreement = self.commissions.set_index('Loan Officer').loc[team_lead]
            commission_bps = com_agreement['Commission BPS']
            max_commission = com_agreement['Maximum']

            lo_commission = row['Amount'] * (commission_bps / 10000)
            if lo_commission <= max_commission:
                return lo_commission + row['GBR']
            else:
                return max_commission + row['GBR']
        else:
            return row['GBR - Adjusted Dollars']

    def calc_net_income(self, row):
        col_name = 'GBR - Adjusted Dollars' if hasattr(row, 'GBR - Adjusted Dollars') else 'GBR'
        try:
            com_agreement = self.commissions.set_index('Loan Officer').loc[row['Loan Officer']]
            commission_bps = com_agreement['Commission BPS']
            max_commission = com_agreement['Maximum']
        except:
            commission_bps = 100
            max_commission = 6000

        lo_commission = row['Amount'] * (commission_bps / 10000)
        try:
            if lo_commission <= max_commission:
                return row[col_name] - lo_commission
            else:
                return row[col_name] - max_commission
        except:
            print('Failed Here', f'\nlo_commssion: ', lo_commission, f"\nlo: {row['Loan Officer']}")
            return 0

    @staticmethod
    def rename_on_key(check_value: str, mutation_table: pd.DataFrame, mutation_key: str, remove_white_space=False):
        check_value = check_value.replace(' ', '') if remove_white_space else check_value
        if check_value in list(mutation_table[mutation_key]):
            return mutation_table.set_index(mutation_key).loc[check_value][0]
        else:
            return check_value

    @staticmethod
    def calc_revenue_from_margin(margin_bps, gross_volume):
        bps_fac = 10000  # Ten thousand
        return (margin_bps / bps_fac) * gross_volume

    @staticmethod
    def calc_margin_from_revenue(revenue, gross_volume):
        bps_fac = 10000  # Ten thousand
        return (revenue / gross_volume) * bps_fac
