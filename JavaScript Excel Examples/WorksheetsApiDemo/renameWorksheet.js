(function (){
    var appId = "your id";
    var appKey = "your key"; 
    
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);

    var name = "RenameWorksheet.xlsx";
    var opts = {
        folder: "input"
    };
    var sheetName = "Sheet2";
    var newname = "newSheetName";
    instance.renameWorksheet(name, sheetName, newname, opts, function (error, data, response) {
            if (error) throw error;
        });
})();