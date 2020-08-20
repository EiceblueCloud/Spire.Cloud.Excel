(function (){
    var appId = "your id";
    var appKey = "your key"; 
    
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);

    var name = "ProtectWorksheet.xlsx";
    var sheetName = "Sheet3";
    var protectParameter = new SpirecloudExcel.ProtectSheetParameter();
    protectParameter.password = 123;
    protectParameter.protectionType = "scenarios";
    protectParameter.allowSelectingLockedCell = true;
    protectParameter.allowDeletingColumn = true;
    protectParameter.allowFormattingRow = true;
    var opts = {
        folder: "input"
    };
    instance.protectWorksheet(name, sheetName, protectParameter, opts, function (error, data, response) {
            if (error) throw error;
        });
})();