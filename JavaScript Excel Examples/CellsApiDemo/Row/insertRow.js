(function (){
    var appId = "your id";
    var appKey = "your key"; 
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel= require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.CellsApi(apiClient);

    var name = "Rows.xlsx";
    var sheetName = "Sheet1";
    var rowIndex = 3;
    var opts = { folder: "input"};
    instance.insertRow(name, sheetName, rowIndex, opts,
        function (error, data, response) {
            if (error) throw error;
        });
})();