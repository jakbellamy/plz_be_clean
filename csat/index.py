import pandas as pd
import numpy as np
# import bamboolib as bam
from wordcloud import WordCloud, STOPWORDS
import matplotlib.pyplot as plt
from openpyxl import load_workbook

pd.options.mode.chained_assignment = None

TARGET_MONTH = pd.to_datetime('2021-11-01')

sample_surveys = './sample/Customer Satisfaction Surveys (3099) (7).xlsx'
sample_fundings = './sample/Funding By Referral Source (2006) (10).xlsx'
sample_detailed = './sample/Branch Production Details (3073) (7).xlsx'

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
    df['Loan Number'] = df['Loan Number'].apply(pd.to_numeric)

    df['Team'] = df['Processor'].apply(assign_team)
    df.set_index('Loan Number', inplace=True)
    return df


def read_fundings(file):
    df = pd.read_excel(sample_fundings, sheet_name='Details', skiprows=4)
    df.columns = [col.split('\n')[0] for col in df.columns]
    df['Loan Number'] = df['Loan Number'].apply(pd.to_numeric)
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
    df['Loan Number'] = df['Loan Number'].apply(str)

    return df


def calculate_csat_score(series, floor_value=5):
    non_zero_scores = list((series.apply(lambda x: pd.to_numeric(x, errors='coerce'))
                                  .apply(lambda x: x if x > 0 else np.NaN)
                                  .dropna()))
    greater_or_equal_to_floor = len(
        [n for n in non_zero_scores if n > (floor_value - 1)])
    raw_score = greater_or_equal_to_floor / len(non_zero_scores)
    return round(raw_score * 100, 2)


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


########################################## CTC PART ##########################################


def run_app_to_ctc_report(file):
    df = (pd.read_excel(file, sheet_name='RAW DATA', skiprows=4)
            .dropna(subset=['Application Date', 'Clear To Close']))

    app_dates = [d.date()
                 for d in df['Application Date'].apply(pd.to_datetime)]
    ctc_dates = [d.date() for d in df['Clear To Close'].apply(pd.to_datetime)]
    df['App To CTC'] = np.busday_count(app_dates, ctc_dates)

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

    report_names = [team + ' 20211115' for team in teams.keys()]

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