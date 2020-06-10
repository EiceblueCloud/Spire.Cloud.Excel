import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "AutofitRow.xlsx"
password = ""
storage = ""
folder = "ExcelDocument"
sheet_name = "Sheet1"
row_index = 17  # Range value: 1~1048576
first_column = 1
last_column = 1
api.autofit_row(name, sheet_name, row_index, first_column, last_column, folder=folder, storage=storage)