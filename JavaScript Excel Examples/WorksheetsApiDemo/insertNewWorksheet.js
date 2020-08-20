(function (){
    var appId = "your id";
    var appKey = "your key"; 
    
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);

    var name = "InsertNewWorksheet.xlsx";
    var sheettype = "NormalWorksheet"; //Available values: NormalWorksheet, Chart
    var index = 2;
    var newsheetname = "SheetInsert";
    var opts = {
        folder: "input"
    };
    instance.insertNewWorksheet(name, newsheetname, index, sheettype, opts, function (error, data, response) {
            if (error) throw error;
        });
})();