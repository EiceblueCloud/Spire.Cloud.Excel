(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.CellsApi(apiClient);

    var name= "test1.xlsx";
    var worksheet = "Sheet1";
    var sheetName = "Sheet2";
    var destCellName = "C8";
    var opts = { folder: "input",cellname: "D8" };
    instance.copyCell(name, destCellName, sheetName, worksheet, opts, 
        function (error, data, response) {
            if (error) throw error;
        });
})();