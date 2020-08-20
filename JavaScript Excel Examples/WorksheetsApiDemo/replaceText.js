(function (){
    var appId = "your id";
    var appKey = "your key"; 
    
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);

    var name = "ReplaceText.xlsx";
    var opts = {
        folder: "input",
        matchCase: true
    };
    var sheetName = "Sheet4";
    var oldValue = "North America";
    var newValue = "北美";
    instance.replaceText(name, sheetName, oldValue, newValue, opts, function (error, data, response) {
            if (error) throw error;
        });
})();