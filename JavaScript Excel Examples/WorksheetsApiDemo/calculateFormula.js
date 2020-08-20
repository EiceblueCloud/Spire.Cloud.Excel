(function (){
    var appId = "your id";
    var appKey = "your key"; 
    
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);

    var name = "GetCalculateFormula_SUM.xlsx";
    var sheetName = "Sheet4";
    var formula = "=SUM(D8:D9)";
    var opts = {
        folder: "input"
    };
    instance.calculateFormula(name, sheetName, formula, opts, function (error, data, response) {
            if (error) throw error;
        });
})();