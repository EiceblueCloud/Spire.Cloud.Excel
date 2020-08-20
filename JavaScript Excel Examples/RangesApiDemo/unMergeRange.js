(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.RangesApi(apiClient);
    
    var name = "unmergeRange.xlsx";
    var opts = { folder: "input" };
    var sheetName = "Sheet1";
    var range = new SpirecloudExcel.Range();
    //The start value is "1"
    range.FirstRow = 3;
    //The start value is "1"
    range.FirstColumn = 3;
    range.ColumnCount = 8;
    range.RowCount = 5;
    instance.unmergeRange(name, sheetName, range, opts,
        function (error, data, response) {
        if (error) throw error;
    });
})();