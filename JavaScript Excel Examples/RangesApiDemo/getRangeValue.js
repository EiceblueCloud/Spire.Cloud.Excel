(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.RangesApi(apiClient);
    
    var name = "getRangeValue.xlsx";
    var sheetName = "Sheet1";
    var opts = { folder: "input", namerange: "A1:A3" };
    //The opts also can be set like below
    //var opts = { folder: "input", firstRow: 1, firstColumn: 6, rowCount: 2, columnCount: 2 };
    instance.getRangeValue(name, sheetName, opts,
        function (error, data, response) {
        if (error) throw error;
    });
})();