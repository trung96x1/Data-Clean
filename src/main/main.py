# Build exe file : pyinstaller --onefile --noconsole .\main_selenium.py
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from deep_translator import GoogleTranslator
from fuzzywuzzy import fuzz
from datetime import datetime
import os
from Aplication import Application as App
from Setting import Setting
from Logger import Logger


from Define import *

table = None
listExporter = []
listImporter = []
setting = None
logger = Logger()

class Table:
    def __init__(self, fileName, sheetName):
        self.workbook = load_workbook(fileName)
        self.worksheet = self.workbook[sheetName]
        self.numcol = self.worksheet.max_column
        self.numrow = self.worksheet.max_row
    
    def addColumnToEnd(self, colName):
        self.worksheet.cell(row=1, column=self.numcol + 1).value = colName
        self.numcol += 1
        columLetter = get_column_letter(self.numcol)
        self.worksheet.column_dimensions[columLetter].width = 20
        return self.numcol
    
    def getCellValue(self, rowIndex, colIndex):
        if rowIndex > self.numrow or colIndex > self.numcol:
            return None
        cell = self.worksheet.cell(row=rowIndex, column=colIndex)
        if cell.value == None:
            return None
        return cell.value
    
    def setCellValue(self, rowIndex, colIndex, value):
        if rowIndex > self.numrow or colIndex > self.numcol:
            return False
        cell = self.worksheet.cell(row=rowIndex, column=colIndex)
        cell.value = value
        return True
    
    def setCellColor(self, rowIndex, colIndex, color):
        if rowIndex > self.numrow or colIndex > self.numcol:
            return False
        cell = self.worksheet.cell(row=rowIndex, column=colIndex)
        cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        return True
    
    def fillColumColor(self, colIndex, color):
        for row in range(2, self.numrow + 1):
            self.setCellColor(row, colIndex, color)

    def findColumIndex(self, colName):
        for col in range(1, self.numcol + 1):
            if self.worksheet.cell(row=1, column=col).value == colName:
                return col
        return -1

    def save(self, fileName):
        self.workbook.save(fileName)

def setCountry(row, exportCountryIndex, importCountryIndex):
    global table
    if table == None or exportCountryIndex <= 0 or importCountryIndex <= 0 or table.numrow <= 1:
        return False
    
    datasetIndex = table.findColumIndex(DATASET_COLUMN)
    dataValue = table.getCellValue(row, datasetIndex)
    if "Export" in dataValue: 
        table.setCellValue(row, exportCountryIndex, dataValue.split("(Export)")[0].strip())

        destCountryIndex = table.findColumIndex(DESTINATION_COUNTRY_COLUMN)
        destCountry = table.getCellValue(row, destCountryIndex)
        if destCountry == None or destCountry == "":
            table.setCellValue(row, importCountryIndex, NA)
            table.setCellColor(row, importCountryIndex, RED_CODE)
        else :
            table.setCellValue(row, importCountryIndex, destCountry)
    elif "Import" in dataValue: 
        table.setCellValue(row, importCountryIndex, dataValue.split("(Import)")[0].strip())

        originCountryIndex = table.findColumIndex(ORIGIN_COUNTRY_COLUMN)
        originCountry = table.getCellValue(row, originCountryIndex)
        if originCountry == None or originCountry == "":
            table.setCellValue(row, exportCountryIndex, NA)
            table.setCellColor(row, exportCountryIndex, RED_CODE)
        else :
            table.setCellValue(row, exportCountryIndex, originCountry)
    else :
        table.setCellValue(row, importCountryIndex, NA)
        table.setCellValue(row, exportCountryIndex, NA)
        table.setCellColor(row, exportCountryIndex, RED_CODE)
        table.setCellColor(row, importCountryIndex, RED_CODE)
    return True

def setProduct(row, productIndex):
    global setting 
    global table

    if table == None or productIndex <= 0 or table.numrow <= 1:
        return False
    
    descriptionIndex = table.findColumIndex(DESCRIPTION_COLUMN)
    descriptionValue = table.getCellValue(row, descriptionIndex)
    if descriptionValue == None or descriptionValue == "":
        table.setCellValue(row, productIndex, NA)
        table.setCellColor(row, productIndex, RED_CODE)
        return True
    descriptionValue = descriptionValue.lower()
    listProducts = setting.get("listProduct", [])
    for product in listProducts:  
        if any(keyword in descriptionValue for keyword in product.get("key")):
            table.setCellValue(row, productIndex, product.get("name"))
            return True
    
    engTranslated = GoogleTranslator(source='auto', target='en').translate(descriptionValue)
    logger.logi(f"Translated : {engTranslated}")
    for product in listProducts:  
        if any(keyword in descriptionValue for keyword in product.get("key")):
            table.setCellValue(row, productIndex, product.get("name"))
            return True
    table.setCellValue(row, productIndex, NA)
    table.setCellColor(row, productIndex, RED_CODE)

def setExporter(row, exporterIndex):
    global table, listExporter, setting
    if table == None or exporterIndex <= 0 or table.numrow <= 1:
        return False
    
    listExcludeName = setting.get("listExcludeName", [])
    
    rawExprorterIndex = table.findColumIndex(EXPORTER_COLUMN)
    rawExporterValue = table.getCellValue(row, rawExprorterIndex)
    rawExporterValue = rawExporterValue.strip().lower()
    for item in listExcludeName:
        rawExporterValue = rawExporterValue.replace(item, "")

    for company in listExporter:
        if fuzz.ratio(company, rawExporterValue) >= 80:
            table.setCellValue(row, exporterIndex, company)
            return True
    listExporter.append(rawExporterValue)
    table.setCellValue(row, exporterIndex, rawExporterValue)

def setImporter(row, importerIndex):
    global table, listImporter, setting
    if table == None or importerIndex <= 0 or table.numrow <= 1:
        return False
    
    listExcludeName = setting.get("listExcludeName", [])
    
    rawImporterIndex = table.findColumIndex(IMPORTER_COLUMN)
    rawImporterValue = table.getCellValue(row, rawImporterIndex)
    rawImporterValue = rawImporterValue.strip().lower()
    for item in listExcludeName:
        rawImporterValue = rawImporterValue.replace(item, "")

    for company in listImporter:
        if fuzz.ratio(company, rawImporterValue) >= 70:
            table.setCellValue(row, importerIndex, company)
            return True
    listImporter.append(rawImporterValue)
    table.setCellValue(row, importerIndex, rawImporterValue)

def setUnitPrice(row, unitPriceIndex, quantityIndex):
    global table, setting
    if table == None or unitPriceIndex <= 0 or quantityIndex <= 0 or table.numrow <= 1:
        return False
    
    rawQuantityIndex = table.findColumIndex(QUANTITY_COLUMN)
    rawQuantityValue  = table.getCellValue(row, rawQuantityIndex)

    rawQuantityUnitIndex = table.findColumIndex(QUANTITY_UNIT_COLUMN)
    rawQuantityUnitValue  = table.getCellValue(row, rawQuantityUnitIndex).strip().lower()

    rawValueIndex = table.findColumIndex(VALUE_COLUMN)
    rawValueValue  = table.getCellValue(row, rawValueIndex)

    if not isinstance(rawQuantityValue, (int, float)) or not isinstance(rawValueValue, (int, float)):
        table.setCellValue(row, unitPriceIndex, NA)
        table.setCellColor(row, unitPriceIndex, RED_CODE)
        table.setCellValue(row, quantityIndex, NA)
        table.setCellColor(row, quantityIndex, RED_CODE)
        return False

    quantity = 0
    if rawQuantityUnitValue == None or rawQuantityUnitValue == "":
        quantity = -1
    else:
        weightUnit = setting.get("weightUnit", [])
        isValid = False
        for unit in weightUnit:
            unitExchange = unit.get("exchange")
            unitKeys = unit.get("key")
            for key in unitKeys:
                if rawQuantityUnitValue == key:
                    quantity = rawQuantityValue * unitExchange
                    isValid = True
                    break
            if isValid:
                break
        if not isValid:
            quantity = -1

    # Set the quantity
    if quantity == -1:
        table.setCellValue(row, quantityIndex, NA)
        table.setCellColor(row, quantityIndex, RED_CODE)
    else :
        table.setCellValue(row, quantityIndex, quantity)

    # Set the unit price
    if rawValueValue == None or rawValueValue == "" or quantity == -1:
        table.setCellValue(row, unitPriceIndex, NA)
        table.setCellColor(row, unitPriceIndex, RED_CODE)
    else :
        table.setCellValue(row, unitPriceIndex, round(rawValueValue / quantity, 2))

def setTime(row, monthIndex, yearIndex):
    global table
    if table == None or monthIndex <= 0 or yearIndex <= 0 or table.numrow <= 1:
        return False
    
    rawDateIndex = table.findColumIndex(DATE_COLUMN)
    rawDateValue = table.getCellValue(row, rawDateIndex)

    date_obj = datetime.strptime(rawDateValue, "%Y-%m-%d")
    table.setCellValue(row, monthIndex, date_obj.month)
    table.setCellValue(row, yearIndex, date_obj.year)

class Scenario:
    def __init__(self):
        pass 
    def execute(self, file_path, app):
        global table
        startTime = datetime.now()
        if file_path:
            table = Table(file_path, SHEET_NAME)
        else:
            table = Table(TEST_FILE, SHEET_NAME)
        exportCountryIndex = table.addColumnToEnd(EXPORT_COUNTRY_COLUMN)
        table.fillColumColor(exportCountryIndex, YELLOW_CODE)  # Yellow

        importCountryIndex = table.addColumnToEnd(IMPORT_COUNTRY_COLUMN)
        table.fillColumColor(importCountryIndex, YELLOW_CODE)  # Yellow

        productIndex = table.addColumnToEnd(PRODUCT_COLUMN)
        table.fillColumColor(productIndex, YELLOW_CODE)  # Yellow

        exportIndex = table.addColumnToEnd(EXPORTER2_COLUMN)
        table.fillColumColor(exportIndex, YELLOW_CODE)  # Yellow

        importIdex = table.addColumnToEnd(IMPORTER2_COLUMN)
        table.fillColumColor(importIdex, YELLOW_CODE)  # Yellow

        unitPriceIndex = table.addColumnToEnd(UNIT_PRICE_COLUMN)
        table.fillColumColor(unitPriceIndex, YELLOW_CODE)  # Yellow

        quantityIndex = table.addColumnToEnd(QUANTITY_KG_COLUMN)
        table.fillColumColor(quantityIndex, YELLOW_CODE)  # Yellow

        monthIndex = table.addColumnToEnd(MONTH_COLUMN)
        table.fillColumColor(monthIndex, YELLOW_CODE)  # Yellow

        yearIndex = table.addColumnToEnd(YEAR_COLUMN)
        table.fillColumColor(yearIndex, YELLOW_CODE)  # Yellow

        for row in range(2, table.numrow + 1):
            logger.logi(f"Processing {row}/{table.numrow}")
            app.setProgress(row, table.numrow)
            #1 Set the country
            setCountry(row, exportCountryIndex, importCountryIndex)

            #2 Set the product
            setProduct(row, productIndex)

            #3 Set the exporter
            setExporter(row, exportIndex)

            #4 Set the importer
            setImporter(row, importIdex)

            #5 Set the unit price, and quantity
            setUnitPrice(row, unitPriceIndex, quantityIndex)

            #6 Set the time
            setTime(row, monthIndex, yearIndex)
  
        
        dataFoler = os.path.join(os.getcwd(), "Data")
        if not os.path.exists(dataFoler):
            os.makedirs(dataFoler)
        fileName = os.path.basename(file_path) if file_path else TEST_FILE
        filePath = os.path.join(dataFoler, f"{fileName.split('.')[0]}_result.xlsx")

        table.save(filePath)
        os.startfile(filePath)
        
        endTime = datetime.now()
        executionTime = (endTime - startTime)
        app.setExecuteTime(executionTime)
        logger.logi(f"Execution completed in {executionTime}")

        
def main():
    global setting
    setting = Setting(SETTING_FILE_PATH)

    scenario = Scenario()
    app = App(scenario)
    app.run()

if __name__ == "__main__":
    main()