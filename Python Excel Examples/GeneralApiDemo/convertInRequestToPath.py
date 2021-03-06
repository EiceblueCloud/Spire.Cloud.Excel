import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.general_api import GeneralApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.general_api.GeneralApi(configuration)
format = "Pdf"   #Supported formats: Xlsx/Xls/Xlsb/Ods/Pdf/Xps/Ps/Pcl
file = "D:/inputFile/charts.xlsx"
password = ""
destPath = "/output/charts.pdf"
api.convert_in_request_to_path(format, destPath,file=file, password=password)