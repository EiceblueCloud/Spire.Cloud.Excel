(function (){
    var appId = "your id";
    var appKey = "your key"; 
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel= require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.CellsApi(apiClient);

    var name = "CopyRows.xlsx";
    var sheetName = "Sheet1";
    var sourceRowIndex = 1;
    var destinationRowIndex = 5;
    var rowNumber = 3;
    var opts = { folder: "input"};
    instance.copyRows(name, sheetName, sourceRowIndex, destinationRowIndex, rowNumber, opts,
        function (error, data, response) {
            if (error) throw error;
        });
})();