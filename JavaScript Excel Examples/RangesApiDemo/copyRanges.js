(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.RangesApi(apiClient);
    
    var name = "copyRanges.xlsx";
    var sheetName = "Sheet1";
    var opts = { folder: "input" };
    var rangeOperate = new SpirecloudExcel.RangeCopyRequest();
    rangeOperate.operate = "copystyle";  
    rangeOperate.source = new SpirecloudExcel.Range();
    rangeOperate.source.ColumnCount = 10;
    rangeOperate.source.RowCount = 10;
    rangeOperate.source.FirstColumn = 1;
    rangeOperate.source.FirstRow = 1;
    rangeOperate.target = new SpirecloudExcel.Range();
    rangeOperate.target.ColumnCount = 8;
    rangeOperate.target.RowCount = 5;
    rangeOperate.target.FirstColumn = 1;
    rangeOperate.target.FirstRow = 30;
    instance.copyRanges(name, sheetName, rangeOperate, opts, 
      function (error, data, response) {
        if (error) throw error;
    });
})();