from __future__ import print_function
from mailmerge import MailMerge
from datetime import date

template = "WSGC-Annual Report 2017.docx"

CRLDictionary = CRLSheet

document.merge(

    Name=CRLDictionary["Name"],
    Award=CRLDictionary["Award"],
    Advisor=CRLDictionary["Advisor"],
    Project=CRLDictionary["Project"],
    
    TeamMember1=CRLDictionary["Team Member 1"],
    Status1=CRLDictionary["Status 1"],
    Award1=CRLDictionary["Student Award"],
    
    TeamMember2=CRLDictionary["Team Member 2"],
    Status2=CRLDictionary["Status 2"],
    Award2=CRLDictionary["Student Award"],
    
    TeamMember3=CRLDictionary["Team Member 3"],
    Status3=CRLDictionary["Status 3"],
    Award3=CRLDictionary["Student Award"]
    
    TeamMember4=CRLDictionary["Team Member 4"],
    Status4=CRLDictionary["Status 4"],
    Award4=CRLDictionary["Student Award"],
    
    TeamMember5=CRLDictionary["Team Member 5"],
    Status5=CRLDictionary["Status 5"],
    Award5=CRLDictionary["Student Award"]
    
    TeamMember6=CRLDictionary["Team Member 6"],
    Status6=CRLDictionary["Status 6"],
    Award6=CRLDictionary["Student Award"]
    
    
    Absract=CRLDictionary["Abstract"])
    
document.write('test-output.docx')


