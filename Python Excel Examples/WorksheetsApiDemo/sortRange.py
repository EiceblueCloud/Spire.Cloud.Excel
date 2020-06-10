import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "SortRange.xlsx"
folder = "ExcelDocument";
storage = ""
sheetName = "Sheet4"
cellArea = "A1:G10"
dataSorter = spirecloudexcel.models.DataSorter()
dataSorter.case_sensitive = True
dataSorter.has_headers = True
dataSorter.sort_left_to_right = False
KeyList = list()
KeyList.append(spirecloudexcel.models.SortKey(0, "Ascending", "Values"))
KeyList.append(spirecloudexcel.models.SortKey(1, "Descending", "Values"))
KeyList.append(spirecloudexcel.models.SortKey(2, "Top", "BackgroundColor"))
KeyList.append(spirecloudexcel.models.SortKey(3, "Top", "FontColor"))
dataSorter.key_list = KeyList
api.sort_range(name, sheetName, cellArea, dataSorter, folder=folder, storage=storage)
