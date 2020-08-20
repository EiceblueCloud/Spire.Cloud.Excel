(function (){
    var appId = "your id";
    var appKey = "your key"; 
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel= require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.CellsApi(apiClient);

    var name = "DeleteRow.xlsx";
    var sheetName = "Sheet1";
    var startrow = 2;
    var opts = { folder: "input",totalRows: 2 };
    instance.deleteRows(name, sheetName, startrow, opts,
        function (error, data, response) {
            if (error) throw error;
        });
})();