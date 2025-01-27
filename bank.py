import json as j
import datetime

with open("Bank.json") as json_format_file: 
  d = j.load(json_format_file)

import xml.etree.cElementTree as e

root = e.Element("ENVELOPE")

header = e.SubElement(root,"HEADER")

tallyRequest = e.SubElement(header, "TALLYREQUEST").text = "Import Data"

body = e.SubElement(root, "BODY")

importData = e.SubElement(body, "IMPORTDATA")

requestDesc = e.SubElement(importData, "REQUESTDESC")

reportName = e.SubElement(requestDesc, "REPORTNAME").text = "Vouchers"

staticVariables = e.SubElement(requestDesc, "STATICVARIABLES")

svCurrentCompany = e.SubElement(staticVariables, "SVCURRENTCOMPANY").text = "Nimbus Logistics 2020-21"

requestData = e.SubElement(importData, "REQUESTDATA")

def bill_allocations(billType: str, references, amounts, allLedgerEntries: e.SubElement, sign) -> e.SubElement:
    for index, ref in enumerate(references):
        billAllocations = e.SubElement(allLedgerEntries, "BILLALLOCATIONS.LIST")
        billAllocationsName = e.SubElement(billAllocations, "NAME").text = str(ref)
        billAllocationsBillType = e.SubElement(billAllocations, "BILLTYPE").text = billType
        billAllocationsAmount = e.SubElement(billAllocations, "AMOUNT").text = sign + str(amounts[index])
    return allLedgerEntries

for obj in d["Bank"]:
    tallyMessage = e.SubElement(requestData, "TALLYMESSAGE", {"xmlns:UDF":"TallyUDF"})
    voucher = e.SubElement(tallyMessage, "VOUCHER", {"ACTION":"Create", "VCHTYPE":str(obj["Voucher Type"])})
    voucherTypeName = e.SubElement(voucher, "VOUCHERTYPENAME").text = str(obj["Voucher Type"])
    dt_string = obj["Date"]
    format = "%Y-%m-%dT%H:%M:%S.%fZ"
    dt_object = datetime.datetime.strptime(dt_string, format)
    dt_object = dt_object + datetime.timedelta(1)
    date = e.SubElement(voucher, "DATE").text = dt_object.strftime("%Y%m%d")
    reference = e.SubElement(voucher, "REFERENCE")
    naration = e.SubElement(voucher, "NARRATION").text = obj["Comments"]
    guid = e.SubElement(voucher, "GUID")
    alterID = e.SubElement(voucher, "ALTERID")
    arr = [1,2]
    for n in arr:
        allLedgerEntries = e.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
        removeZeroEntries = e.SubElement(allLedgerEntries, "REMOVEZEROENTRIES").text = "No"
        isDeemedPositive = e.SubElement(allLedgerEntries, "ISDEEMEDPOSITIVE")
        if (n==1):
            isDeemedPositive.text = "Yes" # Indicates if an amount is to be credited or debited. Yes means the amount is to be debited.
            if (int(obj["Amount with Sign"])<0):
                ledgerName = e.SubElement(allLedgerEntries, "LEDGERNAME").text = obj["Ledger 1"]
                amount = e.SubElement(allLedgerEntries, "AMOUNT").text = str(obj["Amount with Sign"])
                if (obj["Allocation in Ledger?"]):
                    if (int(obj["Allocation in Ledger?"])==1):
                        allLedgerEntries = bill_allocations("Agst Ref", obj["Against Reference"], obj["Against Reference Amounts"], allLedgerEntries, "-")
                        allLedgerEntries = bill_allocations("Advance", obj["Advance Reference"], obj["Advance Reference Amounts"], allLedgerEntries, "-")
            else:
                ledgerName = e.SubElement(allLedgerEntries, "LEDGERNAME").text = obj["Ledger 2"]
                amount = e.SubElement(allLedgerEntries, "AMOUNT").text = "-" + str(obj["Amount with Sign"])
                if (obj["Allocation in Ledger?"]):
                    if (int(obj["Allocation in Ledger?"])==2):
                        allLedgerEntries = bill_allocations("Agst Ref", obj["Against Reference"], obj["Against Reference Amounts"], allLedgerEntries, "-")
                        allLedgerEntries = bill_allocations("Advance", obj["Advance Reference"], obj["Advance Reference Amounts"], allLedgerEntries, "-")

        elif(n==2):
            isDeemedPositive.text = "No"
            if (int(obj["Amount with Sign"])<0):
                ledgerName = e.SubElement(allLedgerEntries, "LEDGERNAME").text = obj["Ledger 2"]
                amount = e.SubElement(allLedgerEntries, "AMOUNT").text = str(obj["Amount with Sign"])[1:]
                if (obj["Allocation in Ledger?"]):
                    if (int(obj["Allocation in Ledger?"])==2):
                        allLedgerEntries = bill_allocations("Agst Ref", obj["Against Reference"], obj["Against Reference Amounts"], allLedgerEntries, "")
                        allLedgerEntries = bill_allocations("Advance", obj["Advance Reference"], obj["Advance Reference Amounts"], allLedgerEntries, "")

            else:
                ledgerName = e.SubElement(allLedgerEntries, "LEDGERNAME").text = obj["Ledger 1"]
                amount = e.SubElement(allLedgerEntries, "AMOUNT").text = str(obj["Amount with Sign"])
                if (obj["Allocation in Ledger?"]):
                    if (int(obj["Allocation in Ledger?"])==1):
                        allLedgerEntries = bill_allocations("Agst Ref", obj["Against Reference"], obj["Against Reference Amounts"], allLedgerEntries, "")
                        allLedgerEntries = bill_allocations("Advance", obj["Advance Reference"], obj["Advance Reference Amounts"], allLedgerEntries, "")
            

a = e.ElementTree(root)

a.write("bank.xml")


def pretty_xml(element, indent, newline, level=0):  # Elemnt is passed in Elment class parameters for indentation indent, for wrapping NEWLINE
    if element:  # Determine whether the element has child elements    
        if (element.text is None) or element.text.isspace():  # If there is no element of text content
            element.text = newline + indent * (level + 1)
        else:
            element.text = newline + indent * (level + 1) + element.text.strip() + newline + indent * (level + 1)
            # Else: # here two lines if the Notes removed, Element will start a new line of text
            # element.text = newline + indent * (level + 1) + element.text.strip() + newline + indent * level
    temp = list(element)  # Element will turn into a list
    for subelement in temp:
        if temp.index(subelement) < (len(temp) - 1):  # If it is not the last element of the list, indicating that the next line is the starting level of the same elements, indentation should be consistent
            subelement.tail = newline + indent * (level + 1)
        else:  # If it is the last element of the list, indicating that the next line is the end of the parent element, a small indentation should    
            subelement.tail = newline + indent * level
        pretty_xml(subelement, indent, newline, level=level + 1)  # Sub-elements recursion

pretty_xml(root, '\t', '\n')  # Beautification execution method
a = e.ElementTree(root)
a.write('bank_pretty.xml')