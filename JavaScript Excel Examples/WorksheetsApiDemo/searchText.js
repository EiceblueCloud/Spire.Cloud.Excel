(function (){
    var appId = "your id";
    var appKey = "your key"; 
    
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);

    var name = "SearchText.xlsx";
    var sheetName = "Sheet4";
    var text = "America";
    var opt = {
        folder: "input",
        matchCase: false
    };
    instance.searchText(name, sheetName, text, opt, function (error, data, response) {
            if (error) throw error;
        });
})();