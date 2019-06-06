import xlwings as xw
import os
import pandas as pd
import re

def main():
    '''
    Take a directory stored in my_path.
    Search directory and subdirectories for any .xslx files containing
        file_keyword.
    Search those .xlsx files for a tab containing the keyword
        (different than file_keyword).
    Merge those tabs into one dataframe and save it as an excel file.
        ex: save as "compiled_tabs.xlsx".
        The compilation will save in the same directory as the one
        that we are searching.
    '''
    my_path = r'C:\Users\pzuradzki\Downloads\test_excel_tab_pickup'
    keyword = 'sheet'
    file_keyword = 'test'
    all_df = make_all_df(my_path, keyword, file_keyword)
    all_df.to_excel(os.path.join(my_path, "compiled_tabs.xlsx"))

def make_df(my_path, filename, keyword):
    '''
    Takes a file path, filename (must be .xlsx), and keyword for the Excel tab.
    The function will make a dataframe for the Excel range starting in cell A1 of the tab
    that matches the keyword. For example, if the keyword is 'gaps' then any tab with the word 'gap' will
    be stored in a dataframe. If there are multiple tabs matching the keyword, the dataframe will be over-written
    and probably not work as intended.
    '''
    # join the path and filename together for a full file path
    # ex: filepath = /my_path/filename
    filepath = os.path.join(my_path, filename)

    # initiate excel work book
    wb = xw.Book(filepath)

    # get list of sheet names
    sheetnames = [wb.sheets[n].name for n in range(0, len(wb.sheets))]

    # iterate through each sheet/excel tab
    for sheetname in sheetnames:
        # check if keyword is in the sheetname; if yes, proceed
        if keyword.lower() in sheetname.lower():
            # create an excel range object (XLwings object) starting at A1
            sht = wb.sheets[sheetname]
            rng = sht.range('A1')

            # if current_region.value[0] is a list,
            # range object at A1 is a list of lists. Store header AND data.
            if type(rng.current_region.value[0]) == list:
                header, *data = rng.current_region.value
                df = pd.DataFrame(data, columns=header)
            # if .value[0] is a string, that means there is no data
            # only store the headers; dataframe will have no data
            elif type(rng.current_region.value[0]) == str:
                header = rng.current_region.value
                df = pd.DataFrame(columns=header)

            # create a column called source that will contain the filename
            # so we know from which file a tab originates once we consolidate
            df['source'] = filename
            # close the workbook. This step happpens outside the if block,
            # because we can open the file but maybe it doesn't have the keyword
            wb.close()
            return df
    wb.close()
    return None

def file_bool(filename, file_keyword):
    '''
    This checks if filename contains the file_keyword.
    This will be useful to restrict, which excel files we open.
    '''
    return file_keyword.lower() in filename.lower()


def file_xl_bool(filename):
    '''
    This is a regular expression that checks the filename to ensures
    it ends with the .xlsx extension. This ensures we don't open non-Excel
    files and cause an error.
    '''
    my_regex = re.compile(r'.xlsx$')
    match_obj = my_regex.search(filename)
    return match_obj != None

def make_all_df(my_path, keyword, file_keyword):
    '''
    In the function 'make_df', we make one dataframe at a time.
    We want to consolidate the dataframes for each target file that we inspect.
    '''
    # initialize empty datafram that we will append into via pd.concat()
    all_df = pd.DataFrame()

    # loop through each file in given directory
    for foldername, subfolders, filenames in os.walk(my_path):
        for filename in filenames:

            # check if file meets keyword criteria and is a .xlsx
            if file_bool(filename, file_keyword) & file_xl_bool(filename):
                try:
                    filepath = os.path.join(foldername, filename)
                    print(filename)

                    # run make_df to make one dataframe on target sheet/tab
                    df = make_df(foldername, filename, keyword)

                    # consolidate into to all_df by joining all_df to itself +
                    # the new def
                    all_df = pd.concat([all_df,df], sort=False)
                except:
                    print("error")
    return all_df

# if this script is run from the command line,
# then we will automatically call it with defaults in main()
if __name__ == "__main__":
    main()
