(function (){
    var appId = "your id";
    var appKey = "your key"; 
    
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);

    var name = "MoveWorksheet.xlsx";
    var opts = {
        folder: "input"
    };
    var sheetName = "Sheet1";
    var moving = new SpirecloudExcel.WorksheetMovingRequest();
    moving.destinationWorksheet = "sheet4";
    moving.position = "after";
    instance.moveWorksheet(name, sheetName, moving, opts, function (error, data, response) {
            if (error) throw error;
        });
})();