import json as j
import datetime

with open("Bills.json") as json_format_file: 
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

for obj in d["Bills"]:
    tallyMessage = e.SubElement(requestData, "TALLYMESSAGE", {"xmlns:UDF":"TallyUDF"})
    voucher = e.SubElement(tallyMessage, "VOUCHER", {"ACTION":"Create", "VCHTYPE":"Sales"})
    voucherTypeName = e.SubElement(voucher, "VOUCHERTYPENAME").text = "Sales"
    dt_string = obj["Invoice Date"]
    format = "%Y-%m-%dT%H:%M:%S.%fZ"
    dt_object = datetime.datetime.strptime(dt_string, format)
    date = e.SubElement(voucher, "DATE").text = dt_object.strftime("%Y%m%d")
    reference = e.SubElement(voucher, "REFERENCE")
    naration = e.SubElement(voucher, "NARRATION").text = obj["From/To"] + ". " + \
        str(obj["Vehicle Number"]) + ". " + "FM" + str(obj["Challan No"]) + \
        " Comments: " + str(obj["Comments"]) + ". Detention: " + str(obj["Note for Detention Charges"])
    partyName = e.SubElement(voucher, "PARTYNAME").text = obj["Client Name"]
    voucherNumber = e.SubElement(voucher, "VOUCHERNUMBER").text = str(obj["Invoice No"])
    guid = e.SubElement(voucher, "GUID")
    alterID = e.SubElement(voucher, "ALTERID")
    clientAmount = "-" + str(round(obj["Total Bill Amount"],2))
    if (obj["GST (Payable by Us)"]):
        arr = [1,2,3]
        freightForwardingAmount = str(round(round(obj["Total Bill Amount"],2) - round(obj["GST (Payable by Us)"],2),2))
        gstAmount = str(round(obj["GST (Payable by Us)"],2))
        if (float(freightForwardingAmount) + float(gstAmount) != -float(clientAmount)):
            raise Exception("Freight Mismatch Error ", obj["Invoice No"])
    else:
        arr = [1,2]
        freightForwardingAmount = str(round(obj["Total Bill Amount"],2))
        if (float(freightForwardingAmount) != -float(clientAmount)):
            raise Exception("Freight Mismatch Error ", obj["Invoice No"])
    for n in arr:
        allLedgerEntries = e.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
        removeZeroEntries = e.SubElement(allLedgerEntries, "REMOVEZEROENTRIES").text = "No"
        isDeemedPositive = e.SubElement(allLedgerEntries, "ISDEEMEDPOSITIVE")
        if (n==1):
            isDeemedPositive.text = "Yes"
            ledgerName = e.SubElement(allLedgerEntries, "LEDGERNAME").text = obj["Client Name"]
            isPartyLedger = e.SubElement(allLedgerEntries, "ISPARTLEDGER").text = "Yes"
            isLastDeemedPositive = e.SubElement(allLedgerEntries, "ISLASTDEEMEDPOSITIVE").text = "Yes"
            amount = e.SubElement(allLedgerEntries, "AMOUNT").text = clientAmount
            billAllocations = e.SubElement(allLedgerEntries, "BILLALLOCATIONS.LIST")
            billAllocationsName = e.SubElement(billAllocations, "NAME").text = str(obj["Invoice No"])
            billAllocationsBillType = e.SubElement(billAllocations, "BILLTYPE").text = "New Ref"
            billAllocationsAmount = e.SubElement(billAllocations, "AMOUNT").text = clientAmount
        elif(n==2):
            isDeemedPositive.text = "No"
            ledgerName = e.SubElement(allLedgerEntries, "LEDGERNAME").text = "Freight Forwarding (MH)"
            isPartyLedger = e.SubElement(allLedgerEntries, "ISPARTLEDGER").text = "No"
            isLastDeemedPositive = e.SubElement(allLedgerEntries, "ISLASTDEEMEDPOSITIVE").text = "No"
            amount = e.SubElement(allLedgerEntries, "AMOUNT").text = freightForwardingAmount
        else:
            isDeemedPositive.text = "No"
            ledgerName = e.SubElement(allLedgerEntries, "LEDGERNAME").text = "GST Payable"
            isPartyLedger = e.SubElement(allLedgerEntries, "ISPARTLEDGER").text = "No"
            isLastDeemedPositive = e.SubElement(allLedgerEntries, "ISLASTDEEMEDPOSITIVE").text = "No"
            amount = e.SubElement(allLedgerEntries, "AMOUNT").text = gstAmount
            

a = e.ElementTree(root)

a.write("sales.xml")

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
a.write('sales_pretty.xml')