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
    var cellarea = "G2";
    var value = "250";
    var type = "int";
    instance.setRangeValue(name, sheetName, cellarea, value, type, opts, 
        function (error, data, response) {
            if (error) throw error;
        });
})();