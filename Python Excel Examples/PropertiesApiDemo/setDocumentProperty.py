import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.properties_api import PropertiesApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.properties_api.PropertiesApi(configuration)
name = "SetDocumentProperties.xlsx"
password = ""
folder = "Properties"
storage = ""
propertyName = "Author";
properties= spirecloudexcel.models.DocumentProperty("Author", "e-iceblue",True)
api.set_document_property(name, propertyName, _property = properties, password = password, folder = folder, storage = storage)



