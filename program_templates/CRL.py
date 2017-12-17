from __future__ import print_function
from mailmerge import MailMerge
from datetime import date

class CRLTemplate():

    def __init__(self, document, fields):

        self.document = document
        
        self.fields = fields
    
        self.__parse_document()

    def __parse_document(self):

        for entry, recipient in enumerate(self.fields, 1):

            self.document.merge(
        
                Name = recipient["Name"],
                Award = recipient["Award"],
                Adivsor = recipient["Award"],
                Project = recipient["Project"],
                
                TeamMember1 = recipient["TeamMember1"],
                TeamMember2 = recipient["TeamMember2"],
                TeamMember3 = recipient["TeamMember3"],
                TeamMember4 = recipient["TeamMember4"],
                TeamMember5 = recipient["TeamMember5"],
                TeamMember6 = recipient["TeamMember6"],
                                
                Status1 = recipient["Status1"],
                Status2 = recipient["Status2"],
                Status3 = recipient["Status3"],
                Status4 = recipient["Status4"],
                Status5 = recipient["Status5"],
                Status6 = recipient["Status6"],
                                
                Award1 = recipient["StudentAward"],
                Award2 = recipient["StudentAward"],
                Award3 = recipient["StudentAward"],
                Award4 = recipient["StudentAward"],
                Award5 = recipient["StudentAward"],
                Award6 = recipient["StudentAward"],
                                
                Abstract = recipient["Abstract"],
                                
                CongressionalDistrict = recipient["CongressionalDistrict"],
                CongressionalRepresentative = recipient["CongressionalRepresentative"]
                                
                )
                
            document.write('CRL/CRL_%d.docx' % entry)

