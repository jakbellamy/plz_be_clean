from tabula import read_pdf # This library is doing the heavy lifting
import os # built in library to look in my file system.
import pandas as pd # Library i use all the time here.

CONTRACTS_PATH = './lite_contracts' # Lazy so set a variable for the path you wanna search

def read_contract(path): # this method reads the contract, takes the data then spits out a dictionary (in R it's called a hash)
    keys = ['Account', 'Address', 'Contract Signatory', 'Contract Signatory Phone', 'Contract Signatory Email', # I did the work to know what data will be spit out after the logic below, these are the keys to know what each of those values are
            'Contract Date Range']
    data = [] # set an empty list (or array) to push the data from the contract into
    for v in read_pdf(path): # read_pdf gives me a list of dataframes. this isn't ideal. I'm being lazy don't worry about this too much. 
        if list(v.columns)[0] == 'Unnamed: 0': # i use this logic to fix the fact that i'm being lazy don't worry about it
            if len(v.values) > 0:
                data.append(list(v.values)[0])
        else:
            data.append(list(v.columns)[0])
    data = [x.replace('\r', ' ') for x in list(data) if isinstance(x, str)] # This syntax looks weird, it's basically a condensed for loop. This one replaces '\r' artifacts with spaces
    data = [x for x in data if not x == 'Media Company'] # Same syntax. This is basically filtering with a for loop to pull out 'Media Company' which threw off some contracts.
    data = [x.replace('edia Company ', '') for x in data] # Same again. This was to pull out 'edia Company' which sat at the beginning of some addresses. This was all trial and error
    data = dict(zip(keys, data[:-1])) # So I had two lists, keys and data. The zip function puts the lists together, and the dict function instantiates it as a dictionary (or hash). I use [:-1] to go to the second to last item in data because there was always one left over value i didn't want. again lazy
    data['file'] = path.split('/')[-1] # I split the file path string (which was provided as an argument here) on '/' then took the last one to get the file name. Then i set it to a key in the new data dictionary. I did this just for easy look up
    return data # Returns the data object (or hash or dict whatever you wanna call it) when called

contracts = [f'{CONTRACTS_PATH}/{contract}' for contract in os.listdir(CONTRACTS_PATH) if '.pdf' in contract] # this is that list comprehension syntax. Like a for loop that returns a list. I used this to look in my contracts path, find the files then return them as a list of usable file paths.

data = [] # again setting an empty list for filling up
for contract in contracts: # for each contract path in the contract paths list
    data.append(read_contract(contract)) # I run the read_contract method/function and put it in the data list. I shouldn't call every variable data, i'm being lazy

pd.DataFrame(data).to_clipboard(index=False) # The magic of the pandas library DataFrame. I turn that list into a dataframe (like a spreadsheet) and copy it to clipboard. Below i return the dataframe to the python console. Now i can view it and paste it to excel.
pd.DataFrame(data)
