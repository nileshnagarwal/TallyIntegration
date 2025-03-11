## Aim: To convert the google spreadsheet data into Tally XML format for importing into Tally ERP 9

## Input:
- Google SpreadSheet Name: "Nimbus Trip Records <Financial Year eg 2024-25>" or "SSTS Trip Records <Financial Year eg 2024-25>"
- Sheet1: "Challans"
- Sheet2: "Sales"
- Sheet3: "Bank"
- Sheet4: "Ledgers"

### Header Details:
#### Challans:
Challan No
Invoice Number
Challan Date
Client Name
Transporter Name
Vehicle Number
From
To
Return
LR Date
LR No
Ack Received date
Vehicle Category
Vehicle Type
ODC/Normal
Length
Width
Height
Weight
Rate
Challan Freight
Challan Advance
Challan Other Charges
Challan Detention Charges
"Other Charges TDS Non Deductible
(Not included in Adv, Totals and Balance)"
Total Challan Amount (Not Incl TDS Non Deductible Charges)
Own / Hired Vehicle
Comments
Driver Mobile No
Transporter Address
Size
From/To
Narration
Sales Person
Traffic Incharge
Ewaybill Expiry Date
Ewaybill Reminder
TDS Percent Deductible
TDS
TDS Paid
TDS Challan Reference Number
"Transporter 
PAN"
Munshiana
Balance Amount
Challan Finalised

#### Sales:
Challan No
Invoice No
Invoice Date
Client Name
Client Address
From
To
Return
Vehicle Number
Freight Amount
Advance
Other Charges
Detention Charges
GST (Payable by Us)
"Total Bill Amount
(To be Received, Not Deducting Adv)"
Total Bill Amount (Excl GST)
Comments
Balance Receivable
LR No
LR Date
From/To
Note
Was Vehicle Detained?
Detention: No of days
Detention: From Date
Detention: To Date
Detention Charges per day
Rate Detention
Detained At
Note for Detention Charges
Service Tax Payable by
Service Tax Note
Rate
GST Amount
Total Bill Amount (Incl GST)
GSTR1 Applicable?
GSTIN
R-AMT-INK/STAT/DY
Bill Sent On
"Payment Status
(Bill Submitted/ Pending/ Cleared)"
Payment Date
Merged Doc ID - Auto Bills 18-19
Merged Doc URL - Auto Bills 18-19
Link to merged Doc - Auto Bills 18-19
Document Merge Status - Auto Bills 18-19

#### Bank:
Date
Narration
Debit
Credit
Matching Ledger
Confidence
Correct Ledger Names
Match Percentage

#### Ledgers does not have any headers. It has the name of the ledgers as saved in Tally in the first column.

## Output:
- Tally XML file
