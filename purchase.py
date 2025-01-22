import json as j
import datetime

with open("Challans.json") as json_format_file: 
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

for obj in d["Challans"]:
    tallyMessage = e.SubElement(requestData, "TALLYMESSAGE", {"xmlns:UDF":"TallyUDF"})
    voucher = e.SubElement(tallyMessage, "VOUCHER", {"ACTION":"Create", "VCHTYPE":"Purchase"})
    voucherTypeName = e.SubElement(voucher, "VOUCHERTYPENAME").text = "Purchase"
    dt_string = obj["Challan Date"]
    format = "%Y-%m-%dT%H:%M:%S.%fZ"
    dt_object = datetime.datetime.strptime(dt_string, format)
    date = e.SubElement(voucher, "DATE").text = dt_object.strftime("%Y%m%d")
    # .strftime("%d") + obj["Challan Date"].strftime("%m") + \
        # obj["Challan Date"].strftime("%Y")
    # EFFECTIVEDATE is omitted for now
    reference = e.SubElement(voucher, "REFERENCE")
    naration = e.SubElement(voucher, "NARRATION").text = obj["From/To"] + ". " + \
        str(obj["Vehicle Number"]) + ". " + str(obj["Comments"])
    voucherNumber = e.SubElement(voucher, "VOUCHERNUMBER").text = obj["Challan No"]
    guid = e.SubElement(voucher, "GUID")
    alterID = e.SubElement(voucher, "ALTERID")
    for n in [1,2]:
        allLedgerEntries = e.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
        removeZeroEntries = e.SubElement(allLedgerEntries, "REMOVEZEROENTRIES").text = "No"
        isDeemedPositive = e.SubElement(allLedgerEntries, "ISDEEMEDPOSITIVE")
        if (n==1):
            isDeemedPositive.text = "NO"
            ledgerName = e.SubElement(allLedgerEntries, "LEDGERNAME").text = obj["Transport Agent"] + \
                " LH Payable"
            amount = e.SubElement(allLedgerEntries, "AMOUNT").text = str(obj["Total Challan Amount"])
        else:
            isDeemedPositive.text = "YES"
            ledgerName = e.SubElement(allLedgerEntries, "LEDGERNAME").text = obj["Transport Agent"] + \
                " LH Acc"
            amount = e.SubElement(allLedgerEntries, "AMOUNT").text = "-" + str(obj["Total Challan Amount"])


a = e.ElementTree(root)

a.write("purchase.xml")

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
a.write('purchase_pretty.xml')