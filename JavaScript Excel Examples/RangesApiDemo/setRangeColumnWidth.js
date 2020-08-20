(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.RangesApi(apiClient);
    
    var name = "setRangeColumnWidth.xlsx";
    var sheetName = "Sheet1";
    var opts = {folder: "input" };
    var range = new SpirecloudExcel.Range();
    range.FirstRow = 1;
    range.FirstColumn = 1;
    range.ColumnCount = 5;
    range.RowCount = 2;
    var value = 15;
    instance.setRangeColumnWidth(name, sheetName, range, value, opts,
        function (error, data, response) {
        if (error) throw error;
    });
})();