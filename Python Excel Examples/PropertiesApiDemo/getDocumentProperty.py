import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.properties_api import PropertiesApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.properties_api.PropertiesApi(configuration)
name = "GetDocumentProperty.xlsx"
propertyName = "Title"
password = ""
folder = "ExcelDocument"
storage = ""
result=api.get_document_property(name, propertyName, password=password, folder=folder, storage=storage)
