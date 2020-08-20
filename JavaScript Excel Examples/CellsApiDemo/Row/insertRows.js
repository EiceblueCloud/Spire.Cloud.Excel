(function (){
    var appId = "your id";
    var appKey = "your key"; 
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel= require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.CellsApi(apiClient);

    var name = "Rows.xlsx";
    var sheetName = "Sheet1";
    var startrow = 5;
    //Available formatAs types: FormatAsBefore, FormatAsAfter, FormatDefault
    var opts = { folder: "input", totalRows: 2, formatAs: "FormatAsBefore" };
    instance.insertRows(name, sheetName, startrow, opts,
        function (error, data, response) {
            if (error) throw error;
        });
})();