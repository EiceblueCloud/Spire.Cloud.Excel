(function (){
    var appId = "your id";
    var appKey = "your key"; 
    
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);

    var name = "GetNamedRanges.xlsx";
    var opt = {
        folder: "input"
    };
    instance.getNamedRanges(name, opt, function (error, data, response) {
            if (error) throw error;
        });
})();