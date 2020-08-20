(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.CellsApi(apiClient);

    var name= "test1.xlsx";
    var sheetName = "Sheet1";
    var firstIndex = 1;
    var lastIndex = 6;
    var opts = { folder: "input" };
    instance.ungroupColumns(name, sheetName, firstIndex, lastIndex, opts, 
        function (error, data, response) {
            if (error) throw error;
        });
})();