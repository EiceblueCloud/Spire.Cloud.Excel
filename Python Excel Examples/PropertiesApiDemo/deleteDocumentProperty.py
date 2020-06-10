import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.properties_api import PropertiesApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.properties_api.PropertiesApi(configuration)
name = "DeleteDocumentProperty_Title.xlsx"
password = ""
folder = "Properties"
storage = ""
propertyName = "Title"
api.delete_document_property(name, propertyName, password=password, folder=folder, storage=storage)


