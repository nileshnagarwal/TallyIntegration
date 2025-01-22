# Creating a template XML for ledger creation with placeholders for future use

# Define the root element
root = etree.Element("ENVELOPE")

# Define the structure of the Tally XML
header = etree.SubElement(root, "HEADER")
tallyrequest = etree.SubElement(header, "TALLYREQUEST")
tallyrequest.text = "Import Data"

body = etree.SubElement(root, "BODY")
importdata = etree.SubElement(body, "IMPORTDATA")
requestdesc = etree.SubElement(importdata, "REQUESTDESC")
reportname = etree.SubElement(requestdesc, "REPORTNAME")
reportname.text = "All Masters"

staticvariables = etree.SubElement(requestdesc, "STATICVARIABLES")
svcompany = etree.SubElement(staticvariables, "SVCURRENTCOMPANY")
svcompany.text = "Your Company Name"  # Placeholder for company name

requestdata = etree.SubElement(importdata, "REQUESTDATA")

# Template for ledger creation
tallymessage_template = etree.SubElement(requestdata, "TALLYMESSAGE", xmlns="TallyUDF")
ledger_template = etree.SubElement(tallymessage_template, "LEDGER", NAME="Transporter Name LH Payable")  # Placeholder for ledger name

# Include GUID and ALTERID with placeholders
guid_template = etree.SubElement(ledger_template, "GUID")
guid_template.text = "Your GUID Here"  # Placeholder for GUID
parent_template = etree.SubElement(ledger_template, "PARENT")
parent_template.text = "Lorry Hire Payable"
isdeemedpositive_template = etree.SubElement(ledger_template, "ISDEEMEDPOSITIVE")
isdeemedpositive_template.text = "Yes"
alterid_template = etree.SubElement(ledger_template, "ALTERID")
alterid_template.text = "Your ALTERID Here"  # Placeholder for ALTERID
openingbalance_template = etree.SubElement(ledger_template, "OPENINGBALANCE")
openingbalance_template.text = "0"

# Language and name tags with placeholders
languagename_list_template = etree.SubElement(ledger_template, "LANGUAGENAME.LIST")
name_list_template = etree.SubElement(languagename_list_template, "NAME.LIST", TYPE="String")
name_template = etree.SubElement(name_list_template, "NAME")
name_template.text = "Transporter Name LH Payable"  # Placeholder for ledger name
languageid_template = etree.SubElement(languagename_list_template, "LANGUAGEID")
languageid_template.text = "1033"

# Convert the etree to a string
xml_template = etree.tostring(root, pretty_print=True, xml_declaration=True, encoding="UTF-8")

# Save the XML template to a file
output_template_file = "/mnt/data/Tally_Ledger_Import_Template.xml"
with open(output_template_file, "wb") as file:
    file.write(xml_template)

output_template_file
