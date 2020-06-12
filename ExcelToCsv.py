import pandas as pd
import argparse as argparse
import configparser as read_config
import os
import datetime as dt

class ExcelToCsv():

    def __init__(self, fileobject):
        self.class_file_name = fileobject

    def check_sheets(self, filename):
        """
        Reading the excel boook
        Returns:
            Sheet Name
        """
        read_file = pd.ExcelFile(filename, header=None)
        return read_file.sheet_names

    def upload_sheet(self, filename, sheet_name, skip_rows):
        # print(skip_rows)
        file_data = pd.read_excel(filename, sheetname=sheet_name,
                                  skipinitialspace=True, skiprows=int(skip_rows))
        return file_data

    def process_data(self, pd_frame, operation, country_seq, kpi_count, output_file_name):
        new_pd_frame = pd_frame.reset_index()
        # column = new_pd_frame.columns
        # print(column)
        # a.columns = a.iloc[0]
        # c=[]
        datetime = dt.datetime.now().strftime('%Y')
        new_pd_frame.rename(columns={'Unnamed: 0': 'DESC',
                                     "Jan": "JAN-"+datetime,
                                     "Feb": "FEB-"+datetime,
                                     "Mar": "MAR-"+datetime,
                                     "Apr": "APR-"+datetime,
                                     "May": "MAY-"+datetime,
                                     "Jun": "JUNE-"+datetime,
                                     "Jul": "JULY-"+datetime,
                                     "Aug": "AUG-"+datetime,
                                     "Sep": "SEP-"+datetime,
                                     "Nov": "NOV-"+datetime,
                                     "Dec": "DEC-"+datetime,
                                     }, inplace=True)
        new_pd_frame.insert(2, 'Country', 'None')
        count = 0
        seq = 0
        for record_number in new_pd_frame.index:
            new_pd_frame.loc[record_number, 'Country'] = country_seq[seq]
            if count >= (kpi_count-1):
                seq = seq + 1
                count = 0
            else:
                count = count + 1
        values = list(new_pd_frame.columns[3:])
        id_var = list(new_pd_frame.columns[:3])
        unpivot_data = pd.melt(new_pd_frame, id_vars=id_var, value_vars=values)
        required_data = unpivot_data.iloc[:, 1:]
        required_data.rename(columns={'DESC': 'KPI_NAME',
                                      'Country': 'COUNTRY_NAME',
                                      'variable': 'YEAR_MONTH',
                                      'value': 'KPI_VALUE'}, inplace=True)
        required_data.to_csv(output_file_name, index=False)
        
def main(file_name, skip_rows, mutiple_sheet, country_order, kpi_count, output_file_name):
    if os.path.isfile(file_name):
        # print("File is Present")
        call_excel = ExcelToCsv(file_name)
        sheet_list = call_excel.check_sheets(file_name)
        for sheet in sheet_list:
            print("Running for the SheetName {0}".format(sheet))
            df_required = call_excel.upload_sheet(file_name, sheet, skip_rows)
            call_excel.process_data(df_required, 'T', country_order, kpi_count,output_file_name)
            if mutiple_sheet == "Y":
                print("Sheet Iteration Not required")
                break
            else:
                print("Continue for Next Sheet")
    else:
        print("File is not Present")


####################################################################


if __name__ == "__main__":
    """ Parsing the Command Line Argument
    """
    parser = argparse.ArgumentParser(description="This Program will Convert NPS File CSV")
    parser.add_argument('-FileName', type=str, help="Enter the Configuration File Path and Name")
    args = parser.parse_args()
    if not args.FileName:
        parser.error("Enter the Excel File Path and Name")
    else:
        # main(args.FileName)
        file = read_config.ConfigParser()
        file.read(args.FileName)
        options = file.options("NPS")

        if "file_name" in options:
            file_name = file.get("NPS", "FILE_NAME")
        else:
            print("Doesn't Find File Name Section")
            exit(-1)
        if "skip_row" in options:
            skip_rows = int(file.get("NPS", "SKIP_ROW"))
        else:
            print("Doesn't Find Skip Row Section Hence Defaulted to 0")
            skip_rows = 0
        if "country_order" in options:
            country_order = (file.get("NPS", "COUNTRY_ORDER")).split(",")
        else:
            print("Doesn't Find Country Order Section")
            exit(-1)
        if "multiple_sheet" in options:
            multiple_sheet = file.get("NPS", "MULTIPLE_SHEET")
        else:
            print("Doesn't Find multiple_sheet Section Hence Defaulted to 0")
            multiple_sheet = 'N'
        if "kpi_count" in options:
            kpi_count = int(file.get("NPS", "KPI_COUNT"))
        else:
            print("Doesn't Find kpi count Section")
            exit(-1)
        if "output_file_name" in options:
            output_file_name = file.get("NPS", "OUTPUT_FILE_NAME")
        else:
            print("Doesn't Find output file Section")
            exit(-1)

        print("Configuration Mapped")
        # print("Running For Sheet {0}".format(file_name))
        # print("Running Skip Rows {0}".format(skip_rows))
        # print("Country Order {0}".format(country_order))
        main(file_name, skip_rows, multiple_sheet, country_order, kpi_count, output_file_name)