(function (){
    var appId = "your id";
    var appKey = "your key"; 
    
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);

    var name = "GetMergedCell.xlsx";
    var sheetName = "sheet5";
    var mergedCellIndex = 0;
    var opt = {
        folder: "input"
    };
    instance.getMergedCell(name, sheetName, mergedCellIndex, opt,
        function (error, data, response) {
            if (error) throw error;
        });
})();