(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.RangesApi(apiClient);
    
    var name = "setRangeStyle.xlsx";
    var opts = {folder: "input" };
    var sheetName = "Sheet1";
    var rangeOperate = new SpirecloudExcel.RangeSetStyleRequest();
    rangeOperate.Style = new SpirecloudExcel.Style();
    rangeOperate.Style.Font = new SpirecloudExcel.Font();
    rangeOperate.Style.Font.Color = new SpirecloudExcel.Color();
    rangeOperate.Style.Font.Color.A = 255;
    rangeOperate.Style.Font.Color.R = 255;
    rangeOperate.Style.Font.Color.G = 0;
    rangeOperate.Style.Font.Color.B = 0;
    rangeOperate.Style.Font.Underline = "single";
    rangeOperate.Style.Font.Size = 12;
    rangeOperate.Style.Font.Name = "Arial";
    rangeOperate.Style.Font.IsItalic = true;
    rangeOperate.Style.Font.IsStrikeout = true;
    rangeOperate.Style.Font.IsBold = true;

    rangeOperate.Style.borderCollection = [new SpirecloudExcel.Border(), new SpirecloudExcel.Border(), new SpirecloudExcel.Border()];
    rangeOperate.Style.borderCollection[0].LineStyle = "Medium";
    rangeOperate.Style.borderCollection[0].Color = new SpirecloudExcel.Color();
    rangeOperate.Style.borderCollection[0].Color.A = 255;
    rangeOperate.Style.borderCollection[0].Color.R = 255;
    rangeOperate.Style.borderCollection[0].Color.G = 0;
    rangeOperate.Style.borderCollection[0].Color.B = 0;
    rangeOperate.Style.borderCollection[0].BorderType = "EdgeTop";

    rangeOperate.Style.borderCollection[1].LineStyle = "DashDot";
    rangeOperate.Style.borderCollection[1].Color = new SpirecloudExcel.Color();
    rangeOperate.Style.borderCollection[1].Color.A = 255;
    rangeOperate.Style.borderCollection[1].Color.R = 0;
    rangeOperate.Style.borderCollection[1].Color.G = 255;
    rangeOperate.Style.borderCollection[1].Color.B = 0;
    rangeOperate.Style.borderCollection[1].BorderType = "EdgeRight";

    rangeOperate.Style.borderCollection[2].LineStyle = "Thin";
    rangeOperate.Style.borderCollection[2].Color = new SpirecloudExcel.Color();
    rangeOperate.Style.borderCollection[2].Color.A = 0;
    rangeOperate.Style.borderCollection[2].Color.R = 160;
    rangeOperate.Style.borderCollection[2].Color.G = 32;
    rangeOperate.Style.borderCollection[2].Color.B = 240;
    rangeOperate.Style.borderCollection[2].BorderType = "EdgeLeft";

    rangeOperate.Style.HorizontalAlignment = "left";
    rangeOperate.Style.indentLevel = 1;
    rangeOperate.Style.textDirection = "RightToLeft";
    rangeOperate.Style.backgroundColor = new SpirecloudExcel.Color();
    rangeOperate.Style.backgroundColor.A = 0;
    rangeOperate.Style.backgroundColor.R = 255;
    rangeOperate.Style.backgroundColor.G = 255;
    rangeOperate.Style.backgroundColor.B = 0;
    rangeOperate.Style.IsTextWrapped = true;

    var range = new SpirecloudExcel.Range();
    range.FirstColumn = 2;
    range.FirstRow = 1;
    range.ColumnCount = 1;
    range.RowCount = 10;
    rangeOperate.Range = range;
    
    instance.setRangeStyle(name, sheetName, rangeOperate, opts,
        function (error, data, response) {
        if (error) throw error;
    });
})();