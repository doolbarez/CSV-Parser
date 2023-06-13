import openpyxl

csv_files = ['input1.csv', 'input2.csv', 'input3.csv', 'input4.csv']
excel_file = 'output.xlsx'
sheet_names = ['DA', 'IDA-1', 'IDA-2', 'IDA-3']
portfolios_to_include = ['Portfolio;NORDJYSK;AU_400006;60;EUR',
'Portfolio;NORDJYSK;AU_400006;30;EUR',
'Portfolio;MFT;AU_400011;60;EUR',
'Portfolio;MFT;AU_400011;30;EUR',
'Portfolio;MFT;AU_400132;60;EUR',
'Portfolio;MFT;AU_400132;30;EUR',
'Portfolio;DANSKE;AU_400003;60;EUR',
'Portfolio;DANSKE;AU_400003;30;EUR',
'Portfolio;ENERDK;AU_400101;60;EUR',
'Portfolio;ENERDK;AU_400101;30;EUR',
'Portfolio;INCOM;AU_400111;60;EUR',
'Portfolio;INCOM;AU_400111;30;EUR',
'Portfolio;INCOM;AU_500105;60;GBP',
'Portfolio;INCOM;AU_500105;30;GBP',
'Portfolio;NIDHOG;AU_500114;60;GBP',
'Portfolio;NIDHOG;AU_500114;30;GBP',
'Portfolio;ENERGET;AU_500122;60;GBP',
'Portfolio;ENERGET;AU_500122;30;GBP',
'Portfolio;KRIA;AU_400137;60;EUR',
'Portfolio;KRIA;AU_400137;30;EUR',
'Portfolio;RISQ;AU_500125;60;GBP',
'Portfolio;RISQ;AU_500125;30;GBP',
'Portfolio;COPENER;AU_500126;60;GBP',
'Portfolio;COPENER;AU_500126;30;GBP',
]

def process_csv_file(csv_file, portfolios_to_include):
    # Catch any error and continue
    # Read all lines from CSV file
    with open(csv_file, 'r') as f:
        lines = f.readlines()

    # Create a list to store extracted data
    data = []

    # Iterate over lines and extract the required information
    i = 0
    while i < len(lines):
        # Find the index of the next line starting with "Portfolio"
        portfolio_index = next((j for j in range(i, len(lines)) if lines[j].startswith("Portfolio")), None)
        if portfolio_index is None:
            break

        # Check if the current portfolio is in the list of portfolios to include
        portfolio = lines[portfolio_index].strip()
        if portfolios_to_include is not None and portfolio not in portfolios_to_include:
            i = portfolio_index + 1
            continue

        # Check the number of lines until the next "Portfolio"
        num_lines = 0
        for j in range(portfolio_index+1, len(lines)):
            if lines[j].startswith("Portfolio"):
                break
            num_lines += 1
        try:
            if num_lines == 4:
                dates = lines[portfolio_index + 2].strip().split(';')
                values = lines[portfolio_index + 3].strip().split(';')
                values = [v.replace(',', '.') for v in values]
                values = list(map(float, values))
                for date, value in zip(dates, values):
                    data.append((portfolio, date, value,'',''))
            elif num_lines == 8:
                dates = lines[portfolio_index + 2].strip().split(';')
                values = lines[portfolio_index + 3].strip().split(';')
                dates2 = lines[portfolio_index + 6].strip().split(';')
                values2 = lines[portfolio_index + 7].strip().split(';')
                values = [v.replace(',', '.') for v in values]
                values2 = [v.replace(',', '.') for v in values2]
                values = list(map(float, values))
                values2 = list(map(float, values2))
                for date, value, date2, value2 in zip(dates, values,dates2, values2):
                    data.append((portfolio, date, value, date2, value2))
        except:
            pass
        data.append(('', '', '', '', ''))
        i = portfolio_index + num_lines + 5

    return data


# Create a new workbook and sheet in Excel
wb = openpyxl.Workbook()

for csv_file, sheet_name in zip(csv_files, sheet_names):
    data = process_csv_file(csv_file, portfolios_to_include)
    sheet = wb.create_sheet(sheet_name)

    # Write the header row to the sheet
    sheet.append(['Portfolio', 'Date', 'Value', 'Date2', 'Value2'])

    # Write the extracted data to the Excel sheet
    for row in data:
        sheet.append(row)

# Remove the default sheet
wb.remove(wb.worksheets[0])

# Save the workbook to a file
wb.save(excel_file)