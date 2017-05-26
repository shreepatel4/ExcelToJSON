import xlrd
from collections import OrderedDict
import simplejson as json

# Open the workbook and select the first worksheet
wb = xlrd.open_workbook('Bell mapping.xlsx')
sh = wb.sheet_by_index(0)

# List to hold dictionaries
result = {}
result['SubContract'] = []
subcontracts = []

# Iterate through each row in worksheet and fetch values into dict
for rownum in range(1, sh.nrows):
    contract = {'Components':[]}
    component = dict()
    row_values = sh.row_values(rownum)
    vendor = row_values[3]
    vendor = vendor.strip()
    contract['VendorId'] = vendor
    contractNumber = str(row_values[19])
    contractNumber = contractNumber.strip()
    #contractNumber = contractNumber[0:-2]
    contract['ContractNumber'] = contractNumber
    jobNumber = str(row_values[27])
    jobNumber = jobNumber.strip()
    contract['MainJobNumber'] = jobNumber
    firstpart = str(row_values[19])
    firstpart = firstpart.strip()
    firstpart = firstpart[0:-2]
    secondpart = row_values[0]
    contract['SubcontractItemNumber'] = firstpart + '-' + str(int(secondpart))

    #componentdescription "Unit # : + segmentTwo if it isn't null, else use the category description
    segmentTwo = str(row_values[43])
    if len(segmentTwo) > 4:
        segmentTwo = segmentTwo[0:3]
    if len(segmentTwo) == 3:
        segmentTwo = "Unit #:" + segmentTwo
    else:
        segmentTwo =row_values[73]
    contract['ComponentDescription'] = segmentTwo

    headers = [x for x in result.get('SubContract') if x.get('MainJobNumber') == contract.get('MainJobNumber') and\
                                                       x.get('ContractNumber') == contract.get('ContractNumber') and \
                                                       x.get('VendorId') == contract.get('VendorId')]
    if headers:
        item = [x for x in headers[0].get('Components') if x.get('MainJobNumber') == component.get('MainJobNumber') and\
                                                           x.get('ContractNumber') == component.get('ContractNumber') and \
                                                           x.get('VendorId') == component.get('VendorId') and \
                                                           x.get('SubcontractItemNumber') ==\
                                                                   component.get('SubcontractItemNumber')]
        if not item:
            headers[0].get('Components').append(component)
    else:
        contract.get('Components').append(component)
        result.get('SubContract').append(contract)

#Serialize the list of dicts to JSON
print(result)
j = json.dumps(result)

# Write to file
with open('data.json', 'w') as f:
    f.write(j)

