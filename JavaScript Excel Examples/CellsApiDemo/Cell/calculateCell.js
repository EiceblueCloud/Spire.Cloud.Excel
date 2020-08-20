(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.CellsApi(apiClient);

    var name = "CalculateCell.xlsx";
    var opts = { folder: "input" };
    var sheetName = "Sheet1";
    var cellName = "C5";
    instance.calculateCell(name, sheetName, cellName,opts, 
        function (error, data, response) {
            if (error) throw error;
        });
})();