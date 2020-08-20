(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.CellsApi(apiClient);

    var name = "test1.xlsx";
    var opts = { folder: "input"};
    var sheetName = "Sheet1";
    var startRow = 1;
    var startColumn = 1;
    var totalRows = 2;
    var totalColumns = 2;
    instance.mergeCells(name, sheetName, startRow, startColumn, totalRows, totalColumns, opts, 
        function (error, data, response) {
            if (error) throw error;
        });
})();