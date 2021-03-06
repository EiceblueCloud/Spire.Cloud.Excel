(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.CellsApi(apiClient);

    var name= "test1.xlsx";
    var sheetName = "Sheet1";
    var columnIndex = 2;
    var width = 40.0;
    var opts = { folder: "input" };
    instance.setColumnWidth(name, sheetName, columnIndex, width, opts, 
        function (error, data, response) {
            if (error) throw error;
        });
})();