(function (){
    var appId = "your id";
    var appKey = "your key"; 
    
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);

    var name = "UpdateWorksheetZoom_150.xlsx";
    var opts = {
        folder: "input"
    };
    var sheetName = "Sheet2";
    var value = 150;
    instance.updateWorksheetZoom(name, sheetName, value, opts, function (error, data, response) {
            if (error) throw error;
        });
})();