import pandas as pd


def set_clipboard(str):
    pd.io.clipboard.clipboard_set(str)
