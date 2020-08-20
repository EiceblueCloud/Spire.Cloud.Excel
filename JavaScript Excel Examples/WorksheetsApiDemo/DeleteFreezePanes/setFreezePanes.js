(function (){
    var appId = "your id";
    var appKey = "your key"; 
    
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);

    var name = "SetFreezePanes.xlsx";
    var sheetName = "Sheet2";
    var freezedRows = 3;
    var freezedColumns = 3;
    var opts = {
        folder: "input"
    };
    instance.setFreezePanes(name, sheetName, freezedRows, freezedColumns, opts,
        function (error, data, response) {
            if (error) throw error;
        });
})();