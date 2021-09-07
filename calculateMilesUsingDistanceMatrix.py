from utils import *

excelFile = 'Drivers Trip Report From MAY.xlsx'

sheetNamesList = getSheetNames(excelFile)
print(sheetNamesList)
workingSheet = sheetNamesList[len(sheetNamesList)-1]
print(workingSheet)
print('==================================================================')

fileData = readExcel(excelFile, workingSheet)
index = 0
listOfDrivers = {}
dataWithEachDrivenLocation = {}
additionalDataForOutput = []


for i in fileData['Delivery City '][index:]:
    driver = fileData['Driver / TrK # 300'][index:].tolist()[1]
    if isinstance(driver, str):
        driver = driver.strip()
    locationCity = fileData['Location City'][index:].tolist()[1]
    pickupCity = fileData['Pick up City '][index:].tolist()[1]
    deliveryCity = fileData['Delivery City '][index:].tolist()[1]


    if driver not in listOfDrivers and not str(driver).startswith('Driver') and isinstance(driver, str):
        listOfDrivers[driver] = 0
        additionalDataForOutput = []

    if locationCity == 'Location City':
        pass

    elif isinstance(deliveryCity, str) and isinstance(locationCity, str):
        print(driver, '\t\t', locationCity, '\t\t', pickupCity, '\t\t', deliveryCity, '\t\t')
        parkedToPickUp = calculateDistanceUsingDistanceMatrix(locationCity, pickupCity)
        pickUpToDrop = calculateDistanceUsingDistanceMatrix(pickupCity, deliveryCity)
        total = parkedToPickUp + pickUpToDrop
        print('Miles to Pick up:', parkedToPickUp, '| Miles to Drop off:', pickUpToDrop)
        print('Subtotal:', total)
        print('------------------------------------------------------------------------------')

        if driver in listOfDrivers:
            eachLoadDetails = 'Miles to Pick up:', parkedToPickUp, '| Miles to Drop off:', pickUpToDrop
            listOfDrivers[driver] = listOfDrivers[driver] + total
            additionalDataText = eachLoadDetails, '-', pickupCity,  'to', deliveryCity
            additionalDataForOutput.append(additionalDataText)
            dataWithEachDrivenLocation[driver] = additionalDataForOutput

    elif isinstance(deliveryCity, str) and not isinstance(locationCity, str):
        print(driver, '\t\t', locationCity, '\t\t', pickupCity, '\t\t', deliveryCity, '\t\t')
        deliveryCityFrom = fileData['Delivery City '][index - 1:].tolist()[1]
        if isinstance(deliveryCity, str) and isinstance(deliveryCityFrom, str):
            fromDropToDrop = calculateDistanceUsingDistanceMatrix(deliveryCity, deliveryCityFrom)
            total = total + fromDropToDrop
            print(deliveryCityFrom)
            print('*** Miles from delivery city to next delivery city :', fromDropToDrop)
            print('Subtotal:', total)
            print('------------------------------------------------------------------------------')

        if driver in listOfDrivers:
            eachLoadDetails = '*** Miles from delivery city to next delivery city :', fromDropToDrop
            listOfDrivers[driver] = listOfDrivers[driver] + total
            additionalDataText = eachLoadDetails, '-', deliveryCityFrom, 'to', deliveryCity
            additionalDataForOutput.append(additionalDataText)

            dataWithEachDrivenLocation[driver] = additionalDataForOutput

    if index + 3 <= len(fileData['Driver / TrK # 300']):
        index += 1
        total = 0
    else:
        break

with open(str(workingSheet).strip() + '.csv', 'w') as testDataFile:
    writer = csv.writer(testDataFile, delimiter=',')
    writer.writerow(['Driver', 'Total Miles', 'Additional Info'])
    for i in listOfDrivers:
        print(i, listOfDrivers[i])
        # writer.writerow([i, listOfDrivers[i], dataWithEachDrivenLocation[i]])
        writer.writerow([i, listOfDrivers[i], workingSheet])

print(dataWithEachDrivenLocation)
