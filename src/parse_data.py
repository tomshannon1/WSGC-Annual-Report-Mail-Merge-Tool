from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import pandas as pd
import pprint
import sys
sys.path.append("/Users/tomshannon/Documents/GitHub/WSGC-Annual-Report-Mail-Merge-Tool/program_templates")
from CRL import CRLTemplate


class ParseReportData():

    def __init__(self, excelFile, template):

        self.excelFile = excelFile
        
        self.template = template
    
        self.__parseProgram()
        
        self.__parseTemplate()

    def __parseProgram(self):

        self.excelSheets = pd.ExcelFile(self.excelFile)
        
        dataframe = self.excelSheets.parse("CRL")
    
        dictionary = dataframe.to_dict()
        
        self.fields = []
    
        for key, value in dictionary.items():
            
            for key1, value1 in value.items():
                
                if 0 <= key1 < len(self.fields):
                
                    self.fields[key1][key] = value1

                else:
                    self.fields.append({key:value1})

    def __parseTemplate(self):

        document = MailMerge(self.template)

        CRLTemplate(self.template, self.fields)

        #document.merge_pages(self.fields)
        #document.write('test-output-mult-custs.docx')



if __name__ == "__main__":

    ParseReportData("test.xlsx", "test.docx")
