(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.CellsApi(apiClient);

    var name = "test1.xlsx";
    var opts = { folder: "input",type: "String"};
    var sheetName = "Sheet1";
    var cellName = "C20";
    var value = "123";
    instance.setCellValue(name, sheetName, cellName, value, opts, 
        function (error, data, response) {
            if (error) throw error;
        });
})();