(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);
    
    var name = "autofitRow.xlsx";
    var opts = {folder: "input"};
    var sheetName = "Sheet1";
    var rowIndex = 17; 
    var firstColumn = 1;
    var lastColumn = 1;
    instance.autofitRow(name, sheetName, rowIndex, firstColumn, lastColumn, opts,
        function (error, data, response) {
        if (error) throw error;
    });
})();