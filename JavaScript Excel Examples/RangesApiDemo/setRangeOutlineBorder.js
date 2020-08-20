(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.RangesApi(apiClient);
    
    var name = "setRangeOutlineBorder.xlsx";
    var opts = { folder: "input" };
    var sheetName = "Sheet1";
    var rangeOperate = new SpirecloudExcel.RangeSetOutlineBorderRequest();
    //Available values: outline, vertical, horizontal, edgeleft
    rangeOperate.borderEdge = "outline";
    //Available values: Medium, DashDot, MediumDashDotDot, Thin
    rangeOperate.borderStyle = "Medium";
    rangeOperate.borderColor = new SpirecloudExcel.Color();
    rangeOperate.borderColor.A = 255;
    rangeOperate.borderColor.R = 255;
    rangeOperate.borderColor.G = 0;
    rangeOperate.borderColor.B = 0;
    var range = new SpirecloudExcel.Range();
    range.FirstRow = 1;
    range.FirstColumn = 1;
    range.ColumnCount = 10;
    range.RowCount = 5;
    rangeOperate.Range = range;
    instance.setRangeOutlineBorder(name, sheetName, rangeOperate, opts, 
      function (error, data, response) {
        if (error) throw error;
    });
})();