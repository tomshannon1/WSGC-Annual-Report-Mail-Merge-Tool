from __future__ import print_function
from mailmerge import MailMerge
from datetime import date

class WriteDocuments():

    def __init__(self, template, fields):
        
        self.template = template
        
        self.fields = fields
    
        self.__parse_document()

    def __parse_document(self):

        for entry, recipient in enumerate(self.fields, 1):
            
            # Create a document based off denoted template
            document = MailMerge(self.template)
            
            # Note the district as integer for file naming
            district = recipient["CongressionalDistrict"]
            
            # Convert Congressional District as string (must be in string form to be copied to word document)
            recipient["CongressionalDistrict"] = str(recipient["CongressionalDistrict"])
            
            # Merge documents allows you to pass a dictionary instead of going through each parameter settings
            document.merge_pages([recipient])
            
            # Write the document to the corresponding folder based on Congressional District
            document.write('/Users/tomshannon/Documents/GitHub/WSGC-Annual-Report-Mail-Merge-Tool/program_templates/CD%d/CRL_%d.docx' % (district, entry))

