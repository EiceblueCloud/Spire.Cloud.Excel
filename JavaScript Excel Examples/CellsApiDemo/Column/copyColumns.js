(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.CellsApi(apiClient);

    var name= "test1.xlsx";
    var sheetName = "Sheet1";
    var sourceColumnIndex = 1;
    var destinationColumnIndex = 8;
    var columnNumber = 2;
    var opts = { folder: "input", worksheet: "Sheet2" };
    instance.copyColumns(name, sheetName, sourceColumnIndex, destinationColumnIndex, columnNumber, opts, 
        function (error, data, response) {
            if (error) throw error;
        });
})();