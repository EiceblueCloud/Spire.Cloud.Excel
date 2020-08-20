(function (){
    var appId = "your id";
    var appKey = "your key"; 
    
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);

    var name = "ChangeVisibilityWorksheet_true.xlsx";
    var sheetName = "Sheet1";
    var isVisible = true;
    var opts = {
        folder: "input"
    };
    instance.changeVisibilityWorksheet(name, sheetName, isVisible, opts, function (error, data, response) {
            if (error) throw error;
        });
})();