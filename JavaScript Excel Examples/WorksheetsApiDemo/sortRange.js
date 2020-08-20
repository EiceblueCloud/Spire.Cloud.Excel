(function (){
    var appId = "your id";
    var appKey = "your key"; 
    
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);

    var name = "SortRange.xlsx";
    var opts = {
        folder: "input"
    };
    var sheetName = "Sheet4";
    var cellArea = "A1:G10";
    var dataSorter = new SpirecloudExcel.DataSorter();
    dataSorter.caseSensitive = true;
    dataSorter.hasHeaders = true;
    dataSorter.sortLeftToRight = false;
    dataSorter.keyList = [new SpirecloudExcel.SortKey()];
    dataSorter.keyList[0].key = 0;
    dataSorter.keyList[0].sortOrder = "Ascending";
    dataSorter.keyList[0].sortType = "Values";
    dataSorter.keyList1 = [new SpirecloudExcel.SortKey()];
    dataSorter.keyList1[0].key = 1;
    dataSorter.keyList1[0].sortOrder = "Descending";
    dataSorter.keyList1[0].sortType = "Values";
    instance.sortRange(name, sheetName, cellArea, dataSorter, opts, function (error, data, response) {
            if (error) throw error;
        });
})();