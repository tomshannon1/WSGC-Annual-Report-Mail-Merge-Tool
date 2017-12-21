from mailmerge import MailMerge

class WriteDocuments():

    def __init__(self, template, fields, sheet_name):
        
        self.template = template
        
        self.fields = fields
        
        self.sheet_name = sheet_name
    
        self.__parse_document()

    def __parse_document(self):

        for entry, recipient in enumerate(self.fields, start=0):
            
            # Create a document based off denoted template
            document = MailMerge(self.template)
            
            # Note the district as integer for file naming
            district = recipient["CongressionalDistrict"]
            
            # Convert Congressional District as string (must be in string form to be copied to word document)
            recipient["CongressionalDistrict"] = str(recipient["CongressionalDistrict"])
            
            lastName = self.__getLastName(recipient["Name"])
            
            # Merge documents allows you to pass a dictionary instead of going through each parameter settings
            document.merge_pages([recipient])
            
            # Write the document to the corresponding folder based on Congressional District
            document.write('/Users/tomshannon/Documents/GitHub/WSGC-Annual-Report-Mail-Merge-Tool/program_output/CD%d/%s_%s.docx' % (district, lastName, self.sheet_name))

    def __getLastName(self, recipientName):

        # Initialize data to sort by last name
        isLastName = False
        lastName = ""
        
        # If recipient or project does not have a last name then return name
        if " " not in recipientName:
            return recipientName
        
        # Otherwise check name for space and note last name
        for character in recipientName:
            
            # If you pass a space, start noting characters
            if isLastName == True:
                
                # If anything past the space is not in the alphabet, ignore it
                if not character.isalpha():
                    continue
                
                # Add to last name character by character
                lastName += character
            
            # If character is space, then the next character will be the first character in last name
            if character == " ":
                
                isLastName = True

        return lastName


