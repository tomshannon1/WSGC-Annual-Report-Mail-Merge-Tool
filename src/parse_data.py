from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import pandas as pd
import pprint


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
        
        document.merge_pages(self.fields)
        document.write('test-output-mult-custs.docx')

        '''document.merge(
            Name = self.fields[0]["Name"],
            Award = self.fields[0]["Award"],
            Project = self.fields[0]["Project"],
            TeamMember1 = self.fields[0]["TeamMember1"],
            TeamMember2 = self.fields[0]["TeamMember2"],
            TeamMember3 = self.fields[0]["TeamMember3"],
            TeamMember4 = self.fields[0]["TeamMember4"],
            TeamMember5 = self.fields[0]["TeamMember5"],
            TeamMember6 = self.fields[0]["TeamMember6"],
            Status1 = self.fields[0]["Status1"],
            Status2 = self.fields[0]["Status2"],
            Status3 = self.fields[0]["Status3"],
            Status4 = self.fields[0]["Status4"],
            Status5 = self.fields[0]["Status5"],
            Status6 = self.fields[0]["Status6"],
            StudentAward = self.fields[0]["StudentAward"],
            Advisor = self.fields[0]["Advisor"],
            Abstract = self.fields[0]["Abstract"])
        document.write('test-output.docx')'''

if __name__ == "__main__":

    ParseReportData("test.xlsx", "test.docx")
