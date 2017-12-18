from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import pandas as pd
import pprint
import sys
sys.path.append("/Users/tomshannon/Documents/GitHub/WSGC-Annual-Report-Mail-Merge-Tool/program_templates")
from createdocument import WriteDocuments


class ParseReportData():

    def __init__(self, excelFile, template):

        # Excel workbook with all corresponding data about program recipients
        self.excelFile = excelFile
        
        # Templates for each program (dictionary "AOP": "Template (AOP, SBS).docx")
        self.template = template
    
        # Parse the Excel workbook for data on recipients for each program
        self.__parseProgram()

    def __parseProgram(self):

        # Take in Excel workbook with sheets corresponding to all sheets
        self.excelSheets = pd.ExcelFile(self.excelFile)
        
        # Get sheet names and ignore the first sheet (FIXME in excel sheet later)
        self.sheet_names = self.excelSheets.sheet_names
        self.sheet_names.remove("MainSheet")
        
        # Create dataframes for each excel sheet
        self.dataframes = [self.excelSheets.parse(sheet) for sheet in self.sheet_names]

        # Create dictionary objects for each pandas dataframe
        self.dictionary = [frame.to_dict() for frame in self.dataframes]
        
        # Re-arrange dictionary settings to align with document merge
        for sheet_id, dataframe_dictionary in enumerate(self.dictionary, start=0):
            
            self.fields = []
        
            for key, value in dataframe_dictionary.items():
        
                for otherKey, otherValue in value.items():
    
                    if 0 <= otherKey < len(self.fields):
    
                        self.fields[otherKey][key] = otherValue
    
                    else:
                        self.fields.append({key : otherValue})
        
            # Get corresponding template for mailmerge
            self.template_path = self.template[self.sheet_names[sheet_id]]
            
            # Create document based off template
            self.__parseTemplate()


    def __parseTemplate(self):

        # Populate templated document with data from exceel sheet
        WriteDocuments(self.template_path, self.fields)


if __name__ == "__main__":

    # Add all WSGC programs to the report in dictionary template
    ParseReportData("WSGC_Recipient_Data_Report.xlsx", {"CRL"  : "test.docx"})
