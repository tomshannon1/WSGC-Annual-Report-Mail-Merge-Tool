from __future__ import print_function
from mailmerge import MailMerge
from datetime import date

template = "WSGC-Annual Report 2017.docx"

RIPDictionary = RIPSheet

document.merge(

    Name=RIPDictionary["Name"],
    Award=RIPDictionary["Award"],
    Advisor=RIPDictionary["Project"],
    Project=RIPDictionary["Abstract"],
    Absract=RIPDictionary["Biography"],
    CongressionalDistrict=RIPDictionary["Congressional District"],
    CongressionalRepresentative=RIPDictionary["Congressional Representative"])
    
document.write('RIP-Annual Report.docx')