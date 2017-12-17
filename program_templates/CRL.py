from __future__ import print_function
from mailmerge import MailMerge
from datetime import date

class CRLTemplate():

    def __init__(self, document):

        self.document = document
    
        self.__parse_document()

    def __parse_document(self):


template = "WSGC-Annual Report 2017.docx"

CRLDictionary = CRLSheet

document.merge(

    Name=CRLDictionary["Name"],
    Award=CRLDictionary["Award"],
    Advisor=CRLDictionary["Advisor"],
    Project=CRLDictionary["Project"],
    
    TeamMember1=CRLDictionary["TeamMember 1"],
    Status1=CRLDictionary["Status1"],
    Award1=CRLDictionary["StudentAward"],
    
    TeamMember2=CRLDictionary["TeamMember 2"],
    Status2=CRLDictionary["Status2"],
    Award2=CRLDictionary["StudentAward"],
    
    TeamMember3=CRLDictionary["TeamMember 3"],
    Status3=CRLDictionary["Status3"],
    Award3=CRLDictionary["StudentAward"]
    
    TeamMember4=CRLDictionary["TeamMember4"],
    Status4=CRLDictionary["Status4"],
    Award4=CRLDictionary["StudentAward"],
    
    TeamMember5=CRLDictionary["TeamMember5"],
    Status5=CRLDictionary["Status5"],
    Award5=CRLDictionary["StudentAward"]
    
    TeamMember6=CRLDictionary["TeamMember6"],
    Status6=CRLDictionary["Status6"],
    Award6=CRLDictionary["StudentAward"]
    
    
    Absract=CRLDictionary["Abstract"],
    CongressionalDistrict=CRL
    )
    
document.write('test-output.docx')


