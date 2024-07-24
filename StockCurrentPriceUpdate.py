import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import pandas as pd
from GoogleSheetReader import GoogleSheet

WEBSITE_URL_FOR_STOCK_CURRENT_PRICE = 'https://ticker.finology.in/company/'

# Function to fetch and parse the webpage
def fetch_page_content(url):
    response = requests.get(url)
    
    # Check if the request was successful
    if response.status_code != 200:
        print("Failed to retrieve the webpage")
        return None
    
    # Parse the HTML content using Beautiful Soup
    soup = BeautifulSoup(response.text, 'html.parser')
    return soup

def getCurrentPrice(company_name):
    url = f'{WEBSITE_URL_FOR_STOCK_CURRENT_PRICE}{company_name}'

    print(f"Fetching data from {url}")
    # Initial fetch to simulate a refresh
    soup = fetch_page_content(url)
    if not soup:
        return ''
    
    # Find all div elements with the specified classes
    articles = soup.select('span.d-block.h1.currprice')
    # # Iterate through each article to extract the desired span text
    for article in articles:
        # Find the first span element within the div
        price_span = article.find('span')
        
        # Check if the span element exists and print its text
        if price_span:
            price = price_span.get_text(strip=True)
            return price
    return ''

# def getCompanyCodesList(url):
#     response = requests.get(url)

#     # Check if the request was successful
#     if response.status_code != 200:
#         print("Failed to retrieve the webpage")
#         return None
    
#     # Parse the HTML content using Beautiful Soup
#     soup = BeautifulSoup(response.text, 'html.parser')
#     if not soup:
#         return
    
#     articles = soup.select('table')

#     companyNameWithCodeMapping = {}

#     for i in range(1, len(articles)) :
#         for j in range(1, len(articles[i].select('tr'))) :
#             if len(articles[i].select('tr')[j]) > 2 :
#                 td_tag = articles[i].select('tr')[j].select('td')[1]
#                 a_tag = td_tag.find('a')
#                 anchor_text = a_tag.get_text(strip=True) if a_tag else ''

#                 # Extract additional text from the <td> tag
#                 additional_text = td_tag.get_text(strip=True).replace(anchor_text, '').strip()

#                 companyName = anchor_text + ' ' + additional_text
#                 companyName = companyName.lower()
#                 company_code = articles[i].select('tr')[j].select('td')[0].select('a')[1].get_text(strip=True)
#                 companyNameWithCodeMapping[companyName.strip()] = company_code.strip()
#                 #print(f"Company Name: {company_name} | Company Code: {company_code}")
#             # scrape_website(company_code)
#         #print("---------------------------------------------------")
#     return companyNameWithCodeMapping

# def getFileData(fileName):
#     requiredColumns = ['Row Id', 'Stock Name', 'Transaction Type', 'Stock Status', 'Unit Price(Buy/Sell)', 'Quantity']
#     # Open the CSV file in read mode
#     df = pd.read_excel(fileName, sheet_name='Zerodha Balance Sheet', skiprows=2, usecols=requiredColumns)

#     #print(df)
#     holdingData = df[df['Stock Status'] == 'Hold']

#     # Convert the DataFrame to a list of dictionaries
#     data_list = holdingData.to_dict(orient='records')
#     #print(data_list)
#     return data_list

def updateDataSheetOnServer(sheetInstance, sheetHandle, updatedSheetData):
    sheetRawData = sheetInstance.getWorksheetData(sheetHandle)
    for i, row in enumerate(sheetRawData[3:], start=4):
        if row[0] == updatedSheetData[i-4]['rowId']:
            sheetInstance.updateCell(sheetHandle, i, 20, updatedSheetData[i-4]['unitCurrentPrice'])
            sheetInstance.updateCell(sheetHandle, i, 21, updatedSheetData[i-4]['profitLoss'])

def getFormattedData(data):
    sheetData = []
    for i in range(3, len(data)):
        if data[i][3] != 'Hold':
            continue
        jsonData = {
            "companyName": data[i][1].lower(),
            "unitBuyPrice": float(data[i][6]),
            "unitCurrentPrice": 0,
            "profitLoss": 0,
            "quantity": int(data[i][7]),
            "companyCode": '',
            "rowId": data[i][0],
        }
        sheetData.append(jsonData)
    return sheetData

def getDictionaryOfCompanyCodes(data):
    companyNameWithCodeMapping = {}
    for i in range(1, len(data)):
        companyNameWithCodeMapping[data[i][0].lower()] = data[i][1]
    return companyNameWithCodeMapping

def updateCurrentPriceAndProfitInData(sheetData):
    try:
        print("Fetching current price data from the website")
        for data in sheetData:
            companyCode = data['companyCode']
            if companyCode != '':
                currentPrice = getCurrentPrice(companyCode)
                #print(companyCode + " : " + currentPrice)
                data['unitCurrentPrice'] = float(currentPrice)
                data['profitLoss'] = round((float(currentPrice) - float(data['unitBuyPrice']))*int(data['quantity']), 2)
        
        print('Current price and profit updated in the data')
    except Exception as e:
        print(f"Error updateCurrentPriceAndProfitInData: {e}")
        return
    

def main():
    sheet_name = 'Stock Market'
    zerodha_worksheet_name = 'Zerodha Balance Sheet'
    company_codes_worksheet_name = 'CompanyNameWithCodeMapping'

    try:
        googleSheetInstance = GoogleSheet()
        
        print('getting stock sheet handle')
        stockSheetHandle = googleSheetInstance.getSheetHandle(sheet_name)
        
        if stockSheetHandle != None:
            print('stock sheet handle received')
            
            print('getting zerodha worksheet handle')
            zerodhaWorksheetHandle = googleSheetInstance.getWorksheetHandleByWorksheetName(stockSheetHandle, zerodha_worksheet_name)
            
            if zerodhaWorksheetHandle != None :
                print('zerodha worksheet handle received')
                
                print("Getting worksheet data")
                sheetData = googleSheetInstance.getWorksheetData(zerodhaWorksheetHandle)
                
                if sheetData != None and len(sheetData) > 0 :
                    print('worksheet data received')
                    
                    zerodhaSheetData = getFormattedData(sheetData)
                    print('Stock market data is formatted.')
                
                    if(len(zerodhaSheetData) > 0) :
                        companyCodesWorksheetHandle = googleSheetInstance.getWorksheetHandleByWorksheetName(stockSheetHandle, company_codes_worksheet_name)
                        sheetData = googleSheetInstance.getWorksheetData(companyCodesWorksheetHandle)
                        if sheetData != None and len(sheetData) > 0 :
                            companyNameWithCodeMapping = getDictionaryOfCompanyCodes(sheetData)
                            print('Company codes are extracted from the sheet.')
                            
                            for data in zerodhaSheetData:
                                companyName = data['companyName']
                                if companyNameWithCodeMapping.get(companyName) != None:
                                    #print("Mapping found: "+ companyName + " = " + companyNameWithCodeMapping[companyName])
                                    data['companyCode'] = companyNameWithCodeMapping[companyName]
                                # else:
                                #     print("Mapping not found: "+ companyName)
                            print("Company code with the company name updated in the datasheet")

                            print(zerodhaSheetData)

                            updateCurrentPriceAndProfitInData(zerodhaSheetData)

                            updateDataSheetOnServer(googleSheetInstance, zerodhaWorksheetHandle, zerodhaSheetData)                   
    except Exception as e:
        print(f"Error main(): {e}")

if __name__ == '__main__':
    main()