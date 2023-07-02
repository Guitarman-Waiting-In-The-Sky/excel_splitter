from openpyxl import Workbook
import pandas as pd
from openpyxl.styles import Alignment

'''
The current working version of this will split one Excel report into many based on unique column values.  Once complete, this will
add the option to split the file as separate tabs within the same report.

'''

def remove_illegal_characters(path):
        
        path=str(path)
        illegal_characters='\/<>:"|?*[]'
        for chars in path:
            if chars in illegal_characters:
                path=path.replace(chars, '_')
        return path


def remove_formula_like_characters(string):
        
        # Currently this doesnt do formulas, this remove the = sign from the beginning so as not to cause errors when opening new file.

        if string[0]=='=':
            string=string[1:]
        
        return string 

def split_report_test(split_type, input_report_path, column_number_to_split_by):
    
    df = pd.read_excel(input_report_path)
    selected_col_values=df.iloc[:,int(column_number_to_split_by)-1].unique()

    if str(split_type)=='Separate_Files':

        for unique_items in selected_col_values:

            final_destination_wb = Workbook()
            final_destination_ws = final_destination_wb['Sheet']

            row_counter=2
                
            vectorized_df=df.loc[(df.iloc[:,int(column_number_to_split_by)-1]==unique_items)]

            for i in vectorized_df.itertuples():

                header_counter=1

                vectorized_col=vectorized_df.columns

                for headers in vectorized_df.columns:

                    final_destination_ws.cell(row=1, column=header_counter).value=str(headers)
                    final_destination_ws.cell(row=1, column=header_counter).alignment =Alignment(horizontal='left', vertical='top', text_rotation=0, wrap_text=True, shrink_to_fit=False, indent=0)
                    header_counter=header_counter+1

                col_counter=1

                for cols in range(len(vectorized_col)):
                
                    data=remove_formula_like_characters(str(i[cols+1]))
                    final_destination_ws.cell(row=row_counter, column=col_counter).value=data
                    final_destination_ws.cell(row=row_counter, column=col_counter).alignment =Alignment(horizontal='left', vertical='top', text_rotation=0, wrap_text=True, shrink_to_fit=False, indent=0)
                    col_counter=col_counter+1

                row_counter=row_counter+1

            final_destination_wb.save(f'text_{str(unique_items)}.xlsx')
            final_destination_wb.close()

    elif str(split_type)=='Split_Tabs':

        final_destination_wb = Workbook()
            
        for unique_items in selected_col_values:

            if unique_items not in df.columns:

                sheet_name=remove_illegal_characters(str(unique_items)) # removing forbidden characters from worksheet names
                sheet_name=sheet_name[:30] # Excel has issues with sheet names over 31 characters, this may be problematic if you have 10 different sheet names with the first 30 chars the same
                final_destination_ws = final_destination_wb.create_sheet(sheet_name)

                row_counter=2
                    
                vectorized_df=df.loc[(df.iloc[:,int(column_number_to_split_by)-1]==unique_items)]

                for i in vectorized_df.itertuples():

                    header_counter=1

                    vectorized_col=vectorized_df.columns

                    for headers in vectorized_df.columns:

                        final_destination_ws.cell(row=1, column=header_counter).value=str(headers) 
                        final_destination_ws.cell(row=1, column=header_counter).alignment =Alignment(horizontal='center', vertical='top', text_rotation=0, wrap_text=True, shrink_to_fit=False, indent=0)
                        header_counter=header_counter+1

                    col_counter=1

                    for cols in range(len(vectorized_col)):
                    
                        data=remove_formula_like_characters(str(i[cols+1]))
                        final_destination_ws.cell(row=row_counter, column=col_counter).value=data
                        final_destination_ws.cell(row=row_counter, column=col_counter).alignment =Alignment(horizontal='left', vertical='top', text_rotation=0, wrap_text=True, shrink_to_fit=False, indent=0)
                        col_counter=col_counter+1

                    row_counter=row_counter+1

        del final_destination_wb['Sheet'] # By default xlsx create sheet comes with a tab named sheet.  This gets rid of it.
        final_destination_wb.save('Split_Tab_Report.xlsx')
        final_destination_wb.close()
       
