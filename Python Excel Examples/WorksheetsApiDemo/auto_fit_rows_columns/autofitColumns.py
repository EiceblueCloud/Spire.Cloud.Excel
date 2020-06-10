import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "AutofitColumns.xlsx"
password = ""
storage = ""
folder = "ExcelDocument"
sheet_name = "Sheet1"
first_column = 1
last_column = 5
auto_fitter_options = spirecloudexcel.models.AutoFitterOptions(True, True, True)
first_row = 3
last_row = 10
api.autofit_columns(name, sheet_name, first_column, last_column, auto_fitter_options, first_row=first_row,last_row=last_row, storage=storage, folder=folder)