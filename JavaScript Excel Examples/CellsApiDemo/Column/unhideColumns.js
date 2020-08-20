(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.CellsApi(apiClient);

    var name= "test1.xlsx";
    var sheetName = "Sheet2";
    var startColumn = 1;
    var totalColumns = 10;
    var opts = { folder: "input" };
    instance.unhideColumns(name, sheetName, startColumn, totalColumns, opts, 
        function (error, data, response) {
            if (error) throw error;
        });
})();