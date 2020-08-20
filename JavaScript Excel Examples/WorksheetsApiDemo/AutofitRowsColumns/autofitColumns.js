(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);
    
    var name = "autofitColumns.xlsx";
    var opts = { folder: "input",firstRow: 1, lastRow: 2 };
    var sheetName = "Sheet1";
    var firstColumn = 1;
    var lastColumn = 3;
    var autoFitterOptions = new SpirecloudExcel.AutoFitterOptions();
    autoFitterOptions.AutoFitMergedCellsType = "";
    autoFitterOptions.AutoFitMergedCells = true;
    autoFitterOptions.IgnoreHidden = true;
    autoFitterOptions.OnlyAuto = true;
    instance.autofitColumns(name, sheetName, firstColumn, lastColumn, autoFitterOptions, opts,
        function (error, data, response) {
        if (error) throw error;
    });
})();