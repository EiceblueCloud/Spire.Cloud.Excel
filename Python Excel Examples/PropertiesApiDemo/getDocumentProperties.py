import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.properties_api import PropertiesApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.properties_api.PropertiesApi(configuration)
name = "GetDocumentProperties.xlsx"
password = ""
folder = "ExcelDocument"
storage = ""
result=api.get_document_properties(name, password=password, folder=folder, storage=storage)
