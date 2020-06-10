import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.workbook_api.WorkbookApi(configuration)

name = "EncryptDocument.xlsx"
storage = ""
folder = "ExcelDocument"
encryption = spirecloudexcel.models.WorkbookEncryptionRequest()
encryption.password = "123"
api.encrypt_document(name, encryption, folder=folder, storage=storage)