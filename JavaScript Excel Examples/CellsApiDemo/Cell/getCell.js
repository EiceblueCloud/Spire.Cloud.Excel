(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.CellsApi(apiClient);

    var name = "test1.xlsx";
    var opts = { folder: "input" };
    var sheetName = "Sheet1";
    var cellOrMethodName = "C8";
    instance.getCell(name, sheetName, cellOrMethodName, opts, 
        function (error, data, response) {
            if (error) throw error;
        });
})();