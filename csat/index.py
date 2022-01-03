# USE
# 1. Download Reports 
#   3099: 
#       Date Range = 'Prior 2 Months From Today'
#   2006: 
#       Date Range =  'Prior 90 Days Thru Today'
#   3073:  
#       Date Range = 'Prior 60 Days
#       Activity Type = 'Fundings'
# 2. Leave file names as they are and place files in './MyReports/' folder.
#   - If you wish to remove the old files, you can, but the program will always grab the latest ADDED file
# 3. Run

import glob
import os
import pandas as pd
import numpy as np
from wordcloud import WordCloud, STOPWORDS
import matplotlib.pyplot as plt
from openpyxl import load_workbook

pd.options.mode.chained_assignment = None

TARGET_MONTH = pd.to_datetime('2021-12-01')
TARGET_ENDS = TARGET_MONTH + pd.offsets.MonthBegin(1)
DAY_STAMP = str(pd.to_datetime("today")).split(' ')[0]

def latest_report(report_number):
    report_files = [file for file in glob.glob('./MyReports/*.xlsx') if str(report_number) in file]
    latest_file = max(report_files, key=os.path.getctime)
    return latest_file

sample_surveys = latest_report(3099)
sample_fundings = latest_report(2006)
sample_detailed = latest_report(3073)

teams = {
    'Team POSITIVE': [
        'cherylbosch',
        'coleenkellner',
        'brandonkiefer'
    ],
    'Team FUN': [
        'jeannefigdore',
        'karenpatterson',
        'sonjaharwood'
    ],
    'Team BEST': [
        'lauriebullock',
        'caribaker',
        'debbiemcintyre'
    ]
}

_teams = {
    'Team POSITIVE': [
        'Cheryl Bosch',
        'Coleen Kellner',
        'Brandon Kiefer'
    ],
    'Team FUN': [
        'Jeanne Figdore',
        'Karen Patterson',
        'Sonja Harwood'
    ],
    'Team BEST': [
        'Laurie Bullock',
        'Cari Baker',
        'Debbie McIntyre'
    ]
}

def to_date(date):
    return pd.to_datetime(date).date()


def to_month(month):
    return month.strftime('%b')


def listOfTuples(l1, l2):
    return list(map(lambda x, y: (x, y), l1, l2))


def assign_team(processor, full_names=False):
    if full_names:
        team = [team for team in teams if processor in _teams[team]]
    else:
        team = [team for team in teams if processor in teams[team]]
    if any(team):
        return team[0]
    return ''


def read_processor_details(file):
    df = pd.read_excel(file, skiprows=5, sheet_name='Processor Details')
    df.rename(
        columns={'Processor To Loan Number': 'Loan Number'}, inplace=True)

    selected_processor = ''
    ammended_rows = []
    for _, row in df.iterrows():
        if str.isalpha(row['Loan Number'].replace(' ', '')):
            processor = row['Loan Number'].replace(' ', '')
            selected_processor = processor
        if pd.to_numeric(row['Loan Number'], errors='coerce') > 1000:
            row = row.to_dict()
            row['Processor'] = selected_processor
            ammended_rows.append(row)

    df = pd.DataFrame(ammended_rows)
    df['Loan Number'] = df['Loan Number'].apply(lambda x: str(pd.to_numeric(x)))

    df['Team'] = df['Processor'].apply(assign_team)
    df.set_index('Loan Number', inplace=True)
    return df


def read_lo_details(file):
    df = pd.read_excel(file, skiprows=5, sheet_name='LO Details')
    df.rename(
        columns={'Branch To LO To Loan Number': 'Loan Number'}, inplace=True)

    selected_processor = ''
    ammended_rows = []
    for _, row in df.iterrows():
        if str.isalpha(row['Loan Number'].replace(' ', '')):
            lo = row['Loan Number'].replace(' ', '')
            selected_processor = lo
        if pd.to_numeric(row['Loan Number'], errors='coerce') > 1000:
            row = row.to_dict()
            row['Loan Officer'] = selected_processor
            ammended_rows.append(row)

    df = pd.DataFrame(ammended_rows)
    df['Loan Number'] = df['Loan Number'].apply(lambda x: str(pd.to_numeric(x)))

    df.set_index('Loan Number', inplace=True)
    return df


def read_fundings(file):
    df = pd.read_excel(sample_fundings, sheet_name='Details', skiprows=4)
    df.columns = [col.split('\n')[0] for col in df.columns]
    df['Loan Number'] = df['Loan Number'].apply(lambda x: str(pd.to_numeric(x)))
    df.set_index('Loan Number', inplace=True)
    return df


def load_csat_data(csat_file, fundings_file):
    fundings = read_fundings(sample_fundings)
    csat_surveys = read_processor_details(sample_surveys)
    df = csat_surveys.join(
        fundings[['Funded Month', 'Loan Officer', 'Referral Source', 'Referral Name']])
    df = df.reset_index()
    df = df[['Loan Number', 'Funded Month', 'Processor', 'Team', 'Loan Officer', 'Top Box CSAT Feedback', 'Top Box Recommend Feedback', 'Top Box Repeat Feedback', 'Top Box LOCSAT Feedback', 'Top Box LPCSAT Feedback', 'Top Box Closing Feedback', 'Strength Feedback', 'Recommendation Feedback',
             'Notes Feedback', 'Borrower Last Name', 'Borrower First Name', 'Borrower Email', 'Referral Source', 'Referral Name'] + ['LOS System', 'Top Two CSAT Feedback', 'Top Two Recommend Feedback', 'Top Two Repeat Feedback', 'Top Two LOCSAT Feedback', 'Top Two LPCSAT Feedback', 'Top Two Closing Feedback']]
    df = df.dropna(how='all', axis=1)
    df = df.drop(columns=['LOS System'])

    df['Team'] = df['Processor'].apply(assign_team)
    df['Loan Number'] = df['Loan Number'].apply(lambda x: str(pd.to_numeric(x)))

    return df


def calculate_csat_score(series, floor_value=5, as_string=True):
    non_zero_scores = list((series.apply(lambda x: pd.to_numeric(x, errors='coerce'))
                                  .apply(lambda x: x if x > 0 else np.NaN)
                                  .dropna()))
    greater_or_equal_to_floor = len(
        [n for n in non_zero_scores if n > (floor_value - 1)])
    try:
        raw_score = greater_or_equal_to_floor / len(non_zero_scores)
    except:
        raw_score = 0

    if as_string:
        return round(raw_score * 100, 2)
    else:
        return raw_score


def create_word_cloud(df: 'pd.dataframe', col_name: str):
    text = df[col_name].dropna().values
    wordcloud = WordCloud(
        width=3000,
        height=2000,
        background_color='white',
        stopwords=STOPWORDS).generate(str(text))
    fig = plt.figure(
        figsize=(40, 30),
        facecolor='k',
        edgecolor='white')
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis('off')
    plt.tight_layout(pad=0)
    plt.show()


def run_lo_report(top_box_floor=4):
    df = read_lo_details(sample_surveys)
    fundings = read_fundings(sample_fundings)

    df = df.drop(columns=['Loan Officer']).join(fundings[['Loan Officer']])
    loan_officers = list(set(df['Loan Officer']))
    
    grp = df.groupby('Loan Officer')
    scores = []
    for lo in loan_officers:
        if isinstance(lo, str):
            score = {'Loan Officer': lo}
            score['Score'] = calculate_csat_score(grp.get_group(lo)['Top Box CSAT Feedback'],
                                                floor_value=top_box_floor,
                                                as_string=False)
            scores.append(score)
    return pd.DataFrame(scores)


########################################## CTC PART ##########################################


def run_app_to_ctc_report(file):
    df = (pd.read_excel(file, sheet_name='RAW DATA', skiprows=4)
            .dropna(subset=['Application Date', 'Clear To Close']))

    app_dates = [d.date()
                 for d in df['Application Date'].apply(pd.to_datetime)]
    ctc_dates = [d.date() for d in df['Clear To Close'].apply(pd.to_datetime)]
    holidays=['2021-11-25', '2021-11-26', '2021-12-24', '2021-12-25', '2021-12-31']
    df['App To CTC'] = np.busday_count(app_dates, ctc_dates, holidays=holidays) + 1

    df = df[['Loan Number', 'Loan Processor', 'Loan Officer', 'Last Name',
             'Application Date', 'Clear To Close', 'Funding Date Time', 'App To CTC']]

    df.rename(columns={
        'Last Name': 'Borrower Last Name',
        'Funding Date Time': 'Funding Date'
    }, inplace=True)

    df['Loan Number'] = df['Loan Number'].apply(str)

    date_cols = ['Application Date', 'Clear To Close', 'Funding Date']
    for col in date_cols:
        df[col] = df[col].apply(to_date)

    df['Team'] = df['Loan Processor'].apply(
        lambda x: assign_team(x, full_names=True))

    df = df[df['App To CTC'] < 60]
    df = df[df['Funding Date'].apply(lambda x: pd.to_datetime(x) >= TARGET_MONTH)]
    df = df[df['Funding Date'].apply(lambda x: pd.to_datetime(x) < TARGET_ENDS)]
    df = df[df['Team'].apply(lambda x: x in teams.keys())]



    grp = df.groupby('Team')

    data_frames = []

    results = grp.agg('mean')[['App To CTC']].sort_values(by=['App To CTC'])

    data_frames.append(results)

    for team in teams.keys():
        _df = df.groupby('Team').get_group(team)
        _df.drop(columns='Team', inplace=True)
        data_frames.append(_df)

    return data_frames

##########################################  TO RUN  ##########################################
# team_report = TeamReport('Team POSITIVE')

######################################### GET CSAT SCORES ######################################

# team_reports = []
# for team in teams.keys():
#     team_report =

#################################### RUN THE FULL SCRIPT #########################################




def run():
    # STEP 1. Run the CTC Side
    ctc_df_list = run_app_to_ctc_report(sample_detailed)

    report_names = [f'{team} {DAY_STAMP}' for team in teams.keys()]

    reports = listOfTuples(report_names, ctc_df_list[1:])

    print(ctc_df_list[0])

    for filename, report in reports:
        report.to_excel(f'out/Nov1Test/Nov 1/{filename}.xlsx',
                        index=False, sheet_name='Application To CTC Data')

    print('-----------------------------------------')
    # STEP 2. Add the CSAT Data

    df = load_csat_data(sample_surveys, sample_fundings)

    for i, team in enumerate(teams.keys()):
        sdf = df[(df['Team'] == team) & (df['Funded Month'] == TARGET_MONTH)]
        sdf['Funded Month'] = sdf['Funded Month'].apply(to_month)

        print(team, ' CSAT SCORE: ', calculate_csat_score(
            sdf['Top Box CSAT Feedback']))
        feedback = sdf.dropna(
            subset=['Strength Feedback', 'Recommendation Feedback'], how='all')
        feedback = feedback[['Loan Number', 'Processor', 'Loan Officer', 'Borrower Last Name', 'Funded Month',
                            'Strength Feedback', 'Recommendation Feedback']]

        existing_report_file = 'out/Nov1Test/Nov 1/' + report_names[i] + '.xlsx'
        book = load_workbook(existing_report_file)
        writer = pd.ExcelWriter(existing_report_file, engine='openpyxl')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

        feedback.to_excel(writer, sheet_name='Feedback', index=False)

        under_fours_cols = ['Loan Number', 'Processor', 'Loan Officer', 'Borrower Last Name',
                            'Top Box CSAT Feedback', 'Strength Feedback', 'Recommendation Feedback']
        under_fours = sdf[sdf['Top Box CSAT Feedback'].apply(
            lambda x: 0 < int(x) < 4)][under_fours_cols]

        under_fours.to_excel(writer, sheet_name='CSAT Under 4', index=False)

        ## Dynamicallly adjust the widths of all columns
        ## https://towardsdatascience.com/how-to-auto-adjust-the-width-of-excel-columns-with-pandas-excelwriter-60cee36e175e

        # sheets = [()]

        writer.save()
