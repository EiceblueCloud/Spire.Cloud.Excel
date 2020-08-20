(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);
    
    var name = "getComment.xlsx";
    var sheetName = "Sheet1";
    var cellName = "F9";
    var opt = { folder: "input"};
    instance.getComment(name, sheetName, cellName, opt, 
        function (error, data, response) {
        if (error) throw error;
    });
})();