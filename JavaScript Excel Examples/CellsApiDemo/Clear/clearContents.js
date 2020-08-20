(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.CellsApi(apiClient);

    var name = "test1.xlsx";
    var opts = { folder: "input",startRow: 1, startColumn: 1, endRow: 15, endColumn: 15};
    var sheetName = "Sheet2";
    instance.clearContents(name, sheetName, opts, 
        function (error, data, response) {
            if (error) throw error;
        });
})();