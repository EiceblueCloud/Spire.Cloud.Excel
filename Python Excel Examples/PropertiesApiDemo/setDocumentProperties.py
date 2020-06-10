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

property1 = spirecloudexcel.models.DocumentProperty("Keywords", "Set document properties.", False)
property2 = spirecloudexcel.models.DocumentProperty("Author", "Administrator", False)
property3 = spirecloudexcel.models.DocumentProperty("Company", "E-Iceblue", True)
property4 = spirecloudexcel.models.DocumentProperty("Last saved by", "123456@iCloud.com", True)
property5 = spirecloudexcel.models.DocumentProperty("Status", "true", False)

PropertyList = []
PropertyList.append(property1)
PropertyList.append(property2)
PropertyList.append(property3)
PropertyList.append(property4)
PropertyList.append(property5)
properties = spirecloudexcel.models.DocumentProperties(PropertyList)
api.set_document_properties(name, properties=properties, password=password, folder=folder, storage=storage)