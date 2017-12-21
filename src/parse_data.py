from mailmerge import MailMerge
import pandas as pd
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
            self.sheet_name = self.sheet_names[sheet_id]
            self.__parseTemplate()


    def __parseTemplate(self):

        # Populate templated document with data from exceel sheet
        WriteDocuments(self.template_path, self.fields, self.sheet_name)


if __name__ == "__main__":
    
    template_path = "/Users/tomshannon/Documents/GitHub/WSGC-Annual-Report-Mail-Merge-Tool/program_templates/"
    
    
    # Templates for each WSGC Program
    templates = {"CRL" : "test.docx", "GPP": "template_GPP_IIP_NIP_UGR.docx",
                 "IIP": "template_GPP_IIP_NIP_UGR.docx", "NIP" : "template_GPP_IIP_NIP_UGR.docx",
                 "UGR" : "template_GPP_IIP_NIP_UGR.docx", "AOP" : "template_AOP_HEI_RIP_SIP.docx",
                 "HEI" : "template_AOP_HEI_RIP_SIP.docx", "RIP" : "template_AOP_HEI_RIP_SIP.docx",
                 "SIP" : "template_AOP_HEI_RIP_SIP.docx", "SBS" : "template_SBS_UGS.docx",
                 "SSI" : "template_SSI.docx", "UGS" : "template_SBS_UGS.docx",
                 "EBP" : "template_EBP.docx", "USIP":"template_USIP.docx",
                 "OPP" : "template_AOP_HEI_RIP_SIP.docx"}

    # Add leading absolute path to the template documents
    for key, value in templates.items():
        oldValue = value
        templates[key] = template_path + oldValue

    # Add all WSGC programs to the report in dictionary template
    ParseReportData("WSGC_Recipient_Data_Report.xlsx", templates)
