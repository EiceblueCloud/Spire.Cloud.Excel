import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi
from spirecloudexcel.models import CalculationOptions

appId = "your id"
appKey = "your key"
configuration = ExcelConfiguration(appId, appKey)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)

name = "CalculateCell_1.xlsx"
folder = "/Cells/Cell/"
storage = ""
sheetName = "Sheet1"
cellName = "G6"
options = CalculationOptions()
options.calc_stack_size = 0
options.precision_strategy = ""
options.ignore_error = True
options.recursive = True
api.calculate_cell(name, sheetName, cellName, folder=folder, storage=storage)