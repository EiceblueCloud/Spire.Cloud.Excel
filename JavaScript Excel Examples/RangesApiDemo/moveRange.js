(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.RangesApi(apiClient);
    
    var name = "moveRange.xlsx";
    var opts = {folder: "input" };
    var sheetName = "Sheet1";
    var range = new SpirecloudExcel.Range();
    range.FirstRow = 1;
    range.FirstColumn = 1;
    range.ColumnCount = 1;
    range.RowCount = 1;
    var destRow = 25;
    var destColumn = 3;
    instance.moveRange(name, sheetName, range, destRow, destColumn, opts,
        function (error, data, response) {
        if (error) throw error;
    });
})();