(function (){
    var appId = "your id";
    var appKey = "your key"; 
    
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);

    var name = "UnprotectWorksheet.xlsx";
    var sheetName = "Sheet1";
    var protectParameter = new SpirecloudExcel.ProtectSheetParameter();
    protectParameter.password = 123;
    var opts = {
        folder: "input/"
    };
    instance.unprotectWorksheet(name, sheetName, protectParameter, opts, function (error, data, response) {
            if (error) throw error;
        });
})();