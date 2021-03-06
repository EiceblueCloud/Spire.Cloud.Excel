(function (){
    var appId = "your id";
    var appKey = "your key"; 
    
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);

    var name = "DeleteFreezePanes.xlsx";
    var sheetName = "Sheet5";
    var opts = {
        folder: "input"
    };
    instance.deleteFreezePanes(name, sheetName, opts, 
        function (error, data, response) {
            if (error) throw error;
        });
})();