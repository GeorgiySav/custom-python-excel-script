import sys
import pandas as pd
from openpyxl.utils.cell import column_index_from_string, get_column_letter


def column_string_to_index(letter: str):
    return column_index_from_string(letter) - 1
def column_index_to_string(index: int):
    return get_column_letter(index + 1)


def parse_extraction(filepath, sheetname, column):
    try:
        input_st = pd.read_excel(filepath, sheet_name=sheetname, header=None)
    except Exception as e:
        return "Failed to open the excel file: " + str(e)
    
    try:
        output_wk = pd.ExcelFile(filepath)
    except Exception as e:
        return "Failed to open the output file: " + str(e)

    try:
        col_index = column_string_to_index(column)
    except Exception as e:
        return "Invalid column: " + str(e)

    try:
        header = input_st.iloc[[0]].values.flatten().tolist()
        input_st = input_st[1:].sort_values(by=input_st.columns[col_index], na_position='first', kind="mergesort")
    except Exception as e:
        return "Could not sort excel file, column is likely out of range: " + str(e)

    # skip any empty or nan values 
    index = 0
    while index < input_st.shape[0] and (input_st.iloc[index, col_index] == None or pd.isna(input_st.iloc[index, col_index])):
        index += 1

    while index < input_st.shape[0]:
        start_index = index
        start_name = input_st.iloc[start_index, col_index]
        # excel sheet names can only be at most 30 characters long
        clamped_start_name = start_name[max(0,len(start_name)-31):]
        # remove illegal characters
        illegal_chars = "/\\?*:[]"
        for char in illegal_chars:
            clamped_start_name = clamped_start_name.replace(char, "")

        # iterate until you find a value that is different to the beginning
        index += 1
        while index < input_st.shape[0] and input_st.iloc[index, col_index] == start_name:
            index += 1
        
        # load sheet if it exists, otherwise create a new one with the header
        sheet = pd.DataFrame()
        if clamped_start_name in output_wk.sheet_names:
            sheet = pd.read_excel(output_wk, clamped_start_name, header=None)
        else:
            sheet = pd.concat([sheet, pd.DataFrame([header])])
        
        # append the rows to the sheet
        for i in range(start_index, index):
            row = input_st.iloc[[i]].values.flatten().tolist()
            sheet = pd.concat([sheet, pd.DataFrame([row])], ignore_index=True)
        
        # save the sheet
        try:
            with pd.ExcelWriter(filepath, mode="a", if_sheet_exists="overlay") as writer:
                sheet.to_excel(writer, sheet_name=clamped_start_name, header=False, index=False)
        except Exception as e:
            return "Could not write to output: " + str(e)
        
    return "Executed successfully"

if __name__ == "__main__":
    # reads from the command line arguments
    filepath = str(sys.argv[1])
    sheetname = str(sys.argv[2])
    column = str(sys.argv[3])
    print(parse_extraction(filepath, sheetname, column))