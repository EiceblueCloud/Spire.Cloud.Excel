(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);
    
    var name = "autofitRows.xlsx";
    var opts = {onlyAuto : true ,//The default value of onlyAuto is false
                folder: "input"}; 
    var sheetName = "Sheet1";
    var startRow = 17;
    var endRow = 23;
    instance.autofitRows(name, sheetName, startRow, endRow, opts,
        function (error, data, response) {
        if (error) throw error;
    });
})();