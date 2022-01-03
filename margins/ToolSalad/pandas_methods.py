import pandas as pd
import numpy as np
from datetime import datetime

def sub_search(df, column, like_values):
    """
    Filters a dataframe by masking a column for -> like_value of like_values is a substr of column
    :param df: pd.DataFrame
    :param column: str
    :param like_values: [str**] -OR- str
    :return: pd.DataFrame
    """
    if isinstance(like_values, str):
        like_values = [like_values]
    like_values = [str(x).lower() for x in like_values]

    mask = df[column].apply(lambda x: any(val for val in like_values if val in str(x).lower()))
    return df[mask]


def safe_divide(df, target_columns):
    """
    Safely divides two dataframe columns without throwing an err for div by 0.
    :param df: pd.DataFrame
    :param target_columns: list of TWO columns [str*2]
    :return: pandas series
    """
    return df[target_columns[0]].divide(df[target_columns[1]].apply(lambda x: x if x > 0 else np.nan))


def apply_to_multiple_columns(df, columns, operation, na_value=np.nan):
    """
    Applies a single operation to multiple columns of a pandas dataframe
    :param df: pd.DataFrame
    :param columns: list of target columns [str**]
    :param operation: lambda operation or method to be applied to each column in columns;
    :param na_value: fillna value :: defaults to np.nan
    :return: None (pass)
    """
    for column in columns:
        if column in df.columns:
            df[column] = df[column].fillna(na_value).apply(lambda x: operation(x))
        else:
            pass
    pass


def months_elapsed_now(date):
    return (datetime.now() - pd.to_datetime(date)) / np.timedelta64(1, 'M')


def months_elapsed(start, end):
    return (pd.to_datetime(start) - pd.to_datetime(end)) / np.timedelta64(1, 'M')