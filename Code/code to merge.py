# please install XlswWriter, openpyxl before using the code below
import datetime
import math
import os
import glob
import pandas as pd

# please change the folder_path to the corresponding local folder path on the computer
folder_path = '/Users/yuki/Desktop/olga research/companies'
files = glob.glob(os.path.join(folder_path, '*'))

# Create year arrays and companies' name arrays for 58lines companies
years = []
companyname = []

# This for loop is used to fill in two arrays for 58lines company, which are years and companyname
for file in files:
    # open the worksheet for each company
    worksheet = pd.read_excel(file, header=None, engine='openpyxl')
    worksheet.reset_index(drop=True)

    # check whether the company excel has the same format as the given company
    num_rows = len(worksheet.index)
    if num_rows == 58:
        # put all years for the companies into one dataframe
        selected_row = worksheet.loc[9]
        filtered_year = selected_row[selected_row.notnull()]
        n = len(filtered_year)
        for each_year in filtered_year:
            if isinstance(each_year, int):
                date_int = each_year
                dt_obj = datetime.datetime.fromordinal(datetime.datetime(1900, 1, 1).toordinal() + date_int - 2)
                dt_str = dt_obj.strftime('%Y%m%d%H%M%S')
            else:
                dt_str = each_year.strftime('%Y%m%d%H%M%S')
            yy = dt_str[:4]
            years.append(yy)

        # put all companies' name into one dataframe and repeat the company's name based on the number of years
        value = worksheet.iloc[0, 1]
        for i in range(n):
            companyname.append(value)

# information for 58lines companies
data = {'year': years,
        'company': companyname,
        'Total investments': [],
        'Total reinsurers'' share of tech provisions': [],
        'Total debtors': [],
        'Assets (other)': [],
        'Total assets': [],
        'Capital & surplus': [],
        'Total gross provisions': [],
        'Total creditors': [],
        'Liabilities (other)': [],
        'Total liabilities and surplus': [],
        'Gross premiums written': [],
        'Net premiums written': [],
        'Earned premiums': [],
        'Total underwriting income': [],
        'Total underwriting expenses': [],
        'Balance on combined technical account': [],
        'Net investment income': [],
        'Profit/(loss) before tax': [],
        'Profit/(loss) after tax': [],
        'Net profit/(loss) for the financial year': [],
        'Gross premiums written_life technical account': [],
        'Net premiums written_life technical account': [],
        'Earned premiums_life technical account': [],
        'Total revenue': [],
        'Total expenses': [],
        'Balance on life technical account': [],
        'Gross premiums written_non-life technical account': [],
        'Net premiums written_non-life technical account': [],
        'Earned premiums_non-life technical account': [],
        'Total underwriting expenses_non-life technical account': [],
        'Balance on general technical account': [],
        }

# put all the keys of data into an array
keys_array = list(data.keys())
arr = range(16, 58)

# This loop is used to get the information of rest of rows and fill them into data for 58lines company
for file in files:
    # open the worksheet for each company
    worksheet = pd.read_excel(file, header=None, engine='openpyxl')
    worksheet.reset_index(drop=True)
    df_index = 2

    selected_row = worksheet.loc[9]
    filtered_row = selected_row[selected_row.notnull()]

    # check whether the company excel has the same format as the given company
    num_rows = len(worksheet.index)
    if num_rows == 58:
        # if there is number/n.a. in this row, we create a dataframe and put them into it
        for j in arr:
            row = worksheet.loc[j]
            if isinstance(row[2], str) or not math.isnan(row[2]):
                row_num = row[2:]
                filtered_row = row_num[row.notnull()]
                for m in filtered_row:
                    data[keys_array[df_index]].append(m)
                df_index += 1





# Create year arrays and companies' name arrays for 77lines companies
years_77 = []
companyname_77 = []

# This for loop is used to fill in two arrays for 77lines company, which are years and companyname
for file in files:
    # open the worksheet for each company
    worksheet = pd.read_excel(file, header=None, engine='openpyxl')
    worksheet.reset_index(drop=True)

    # check whether the company excel has the same format as the given company
    num_rows = len(worksheet.index)
    if num_rows == 77 or num_rows == 76:
        # put all years for the companies into one dataframe
        selected_row = worksheet.loc[9]
        filtered_year = selected_row[selected_row.notnull()]
        n = len(filtered_year)
        for each_year in filtered_year:
            if isinstance(each_year, int):
                date_int = each_year
                dt_obj = datetime.datetime.fromordinal(datetime.datetime(1900, 1, 1).toordinal() + date_int - 2)
                dt_str = dt_obj.strftime('%Y%m%d%H%M%S')
            else:
                dt_str = each_year.strftime('%Y%m%d%H%M%S')
            yy = dt_str[:4]
            years_77.append(yy)

        # put all companies' name into one dataframe and repeat the company's name based on the number of years
        value = worksheet.iloc[0, 1]
        for i in range(n):
            companyname_77.append(value)

# information for 77lines companies
data_77 = {'year': years_77,
           'company': companyname_77,
           'Fixed assets': [],
           'Intangible fixed assets': [],
           'Tangible fixed assets': [],
           'Other fixed assets': [],
           'Current assets': [],
           'Stock': [],
           'Debtors': [],
           'Other current assets': [],
           'Cash & cash equivalent': [],
           'Total assets': [],
           'Shareholders funds': [],
           'Capital': [],
           'Other shareholders funds': [],
           'Non-current liabilities': [],
           'Long term debt': [],
           'Other non-current liabilities': [],
           'Provisions': [],
           'Current liabilities': [],
           'Loans': [],
           'Creditors': [],
           'Other current liabilities': [],
           'Total shareh. funds & liab.': [],
           'Working capital': [],
           'Net current assets': [],
           'Enterprise value': [],
           'Number of employees': [],
           'Operating revenue (Turnover)': [],
           'Sales': [],
           'Costs of goods sold': [],
           'Gross profit': [],
           'Other operating expenses': [],
           'Operating P/L [=EBIT]': [],
           'Financial P/L': [],
           'Financial revenue': [],
           'Financial expenses': [],
           'P/L before tax': [],
           'Taxation': [],
           'P/L after tax': [],
           'Extr. and other P/L': [],
           'Extr. and other revenue': [],
           'Extr. and other expenses': [],
           'P/L for period [=Net income]': [],
           'Export revenue': [],
           'Material costs': [],
           'Costs of employees': [],
           'Depreciation & Amortization': [],
           'Other operating items': [],
           'Interest paid': [],
           'Research & Development expenses': [],
           'Cash flow': [],
           'Added value': [],
           'EBITDA': []
           }

# put all the keys of data_77 into an array
keys_array_77 = list(data_77.keys())
arr_77 = range(16, 77)
arr_76 = range(16, 76)

# This loop is used to get the information of rest of rows and fill them into data for 77lines company
for file in files:
    # open the worksheet for each company
    worksheet = pd.read_excel(file, header=None, engine='openpyxl')
    worksheet.reset_index(drop=True)
    df_index = 2

    selected_row = worksheet.loc[9]
    filtered_row = selected_row[selected_row.notnull()]

    # check whether the company excel has 77 or 76 lines
    num_rows = len(worksheet.index)
    if num_rows == 77:
        # if there is number/n.a. in this row, we create a dataframe and put them into it
        for j in arr_77:
            row = worksheet.loc[j]
            if isinstance(row[2], str) or not math.isnan(row[2]):
                row_num = row[2:]
                filtered_row = row_num[row.notnull()]
                for m in filtered_row:
                    data_77[keys_array_77[df_index]].append(m)
                df_index += 1
    elif num_rows == 76:
        # if there is number/n.a. in this row, we create a dataframe and put them into it
        for j in arr_76:
            row = worksheet.loc[j]
            if isinstance(row[2], str) or not math.isnan(row[2]):
                row_num = row[2:]
                filtered_row = row_num[row.notnull()]
                for m in filtered_row:
                    data_77[keys_array_77[df_index]].append(m)
                df_index += 1




# Use data and data_77 to write the merged excel with two sheets
dfyc = pd.DataFrame(data)
dfyc_77 = pd.DataFrame(data_77)
merged = '/Users/yuki/Desktop/olga research/merged_companies.xlsx'
writer = pd.ExcelWriter(merged, engine='xlsxwriter')
dfyc.to_excel(writer, sheet_name='58lines', index=False)
dfyc_77.to_excel(writer, sheet_name='77lines', index=False)
writer.close()

'''# For your convenience, this loop is used to help check how many companies' excel do not have the same format as the
# given one. When I run it, it is 47. This means that 57-47 = 10 companies are not in the same format as the given
# company.
num_not58rows = 0
for file in files:
    df = pd.read_excel(file, header=None, engine='openpyxl')
    # Get the number of rows in the dataframe
    num_rows = len(df.index)
    if not num_rows == 58:
        # print the path of the company excel that does not have the same format
        print(file)
        # print the number of rows of the company to compare it with 58 rows
        print(num_rows)
        num_not58rows += 1
print(num_not58rows)'''




