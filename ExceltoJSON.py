import xlrd
from collections import OrderedDict
import simplejson as json

# Open the workbook and select the first worksheet
wb = xlrd.open_workbook('Bell mapping.xlsx')
sh = wb.sheet_by_index(0)

# List to hold dictionaries
subcontracts = []

# Iterate through each row in worksheet and fetch values into dict
for rownum in range(5):
    updatedSubcontracts = OrderedDict()
    row_values = sh.row_values(rownum)
    vendor = row_values[3]
    vendor = vendor.strip()
    updatedSubcontracts['VendorId'] = vendor
    contractNumber = str(row_values[19])
    contractNumber = contractNumber.strip
    #contractNumber = contractNumber[0:-2]
    updatedSubcontracts['ContractNumber'] = contractNumber
    jobNumber = str(row_values[27])
    jobNumber = jobNumber.strip()
    updatedSubcontracts['MainJobNumber'] = jobNumber
    firstpart = str(row_values[19])
    firstpart = firstpart.strip()
    firstpart = firstpart[0:-2]
    secondpart = row_values[0]
    #if secondpart == float(secondpart):
    #    secondpart = str(int(secondpart))
    #else:
    #    secondpart = str(secondpart)
    #print (secondpart)
    #subcontractstring = firstpart + '-' + secondpart
    #updatedSubcontracts['SubcontractItemNumber'] = subcontractstring

    #componentdescription "Unit # : + segmentTwo if it isn't null, else use the category description
    segmentTwo = str(row_values[43])
    if len(segmentTwo) > 4:
        segmentTwo = segmentTwo[0:3]
    if len(segmentTwo) == 3:
        segmentTwo = "Unit #:" + segmentTwo
    else:
        segmentTwo =row_values[73]
    updatedSubcontracts['ComponentDescription'] = segmentTwo

    subcontracts.append(updatedSubcontracts)
    #Serialize the list of dicts to JSON
#result = {}
#result['SubContract']=subcontracts
    j = json.dumps(subcontracts)

    # Write to file
    with open('data.json', 'w') as f:
        f.write(j)

