

import pandas as pd

def main():
    # Load the Excel file to check its structure
    file_path = 'data.xlsx'
    excel_data = pd.ExcelFile(file_path)

    # Display the sheet names to understand the structure
    print(excel_data.sheet_names)

    av_df = pd.read_excel(file_path, sheet_name='AVDetails')
    edr_df = pd.read_excel(file_path, sheet_name='EDRDetails')
    server_df = pd.read_excel(file_path, sheet_name='ServerDetails')

    av_df['Hostname'] = av_df['Hostname'].str.split('.', n=1).str[0]

    # Add AV? column
    server_df['AV?'] = server_df['HOSTNAME'].isin(av_df['Hostname']).apply(lambda x: 'Y' if x else '')
    server_df['EDR?'] = server_df['HOSTNAME'].isin(edr_df['Hostname']).apply(lambda x: 'Y' if x else '')


    print(server_df.head())
    output_path = 'Updated_DATA.xlsx'
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        server_df.to_excel(writer, sheet_name='ServerDetails', index=False)
        av_df.to_excel(writer, sheet_name='AVDetails', index=False)
        edr_df.to_excel(writer, sheet_name='EDRDetails', index=False)


if __name__ == "__main__":
    main()
