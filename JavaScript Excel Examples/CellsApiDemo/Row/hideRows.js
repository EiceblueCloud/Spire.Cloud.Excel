(function (){
    var appId = "your id";
    var appKey = "your key"; 
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel= require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.CellsApi(apiClient);

    var name = "Rows.xlsx";
    var sheetName = "Sheet1";
    var startrow = 4;
    var totalRows = 2;
    var opts = { folder: "input"};
    instance.hideRows(name, sheetName, startrow, totalRows, opts,
        function (error, data, response) {
            if (error) throw error;
        });
})();