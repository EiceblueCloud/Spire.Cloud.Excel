(function (){
    var appId = "your id";
    var appKey = "your key"; 
    
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);

    var name = "CopyWorksheet_destination.xlsx";
    var sourceSheet = "Sheet2";
    var sheetName = "Sheet4";
    var opts = {
        folder: "input",
        sourceFolder: "input",
        sourceWorkbook: "CopyWorksheet_original.xlsx",
    };
    instance.copyWorksheet(name, sheetName, sourceSheet, opts, function (error, data, response) {
            if (error) throw error;
        });
})();