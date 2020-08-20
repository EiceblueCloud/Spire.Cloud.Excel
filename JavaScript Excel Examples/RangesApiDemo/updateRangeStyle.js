(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.RangesApi(apiClient);
    
    var name = "updateRangeStyle.xlsx";
    var opts = { folder: "input" };
    var sheetName = "Sheet1";
    var range = "C2:D6";
    var style = new SpirecloudExcel.Style();
    style.Font = new SpirecloudExcel.Font();
    style.Font.Color = new SpirecloudExcel.Color();
    style.Font.Color.A = 0;
    style.Font.Color.R = 160;
    style.Font.Color.G = 32;
    style.Font.Color.B = 240;
    style.Font.Underline = "single";
    style.Font.Size = 13;
    style.Font.Name = "Arial";
    style.Font.DoubleSize = 2;
    style.Font.IsItalic = true;
    style.Font.IsStrikeout = true;
    style.Font.IsBold = true;

    style.BorderCollection = [new SpirecloudExcel.Border(), new SpirecloudExcel.Border()];
    style.BorderCollection[0].LineStyle = "Medium";
    style.BorderCollection[0].Color = new SpirecloudExcel.Color();
    style.BorderCollection[0].Color.A = 255;
    style.BorderCollection[0].Color.R = 255;
    style.BorderCollection[0].Color.G = 0;
    style.BorderCollection[0].Color.B = 0;
    style.BorderCollection[0].BorderType = "EdgeTop";
    
    style.BorderCollection[1].LineStyle = "DashDot";
    style.BorderCollection[1].Color = new SpirecloudExcel.Color();
    style.BorderCollection[1].Color.A = 255;
    style.BorderCollection[1].Color.R = 0;
    style.BorderCollection[1].Color.G = 255;
    style.BorderCollection[1].Color.B = 0;
    style.BorderCollection[1].BorderType = "EdgeRight";
    style.IndentLevel = 2;
    style.HorizontalAlignment = "left";
    
    instance.updateRangeStyle(name, sheetName, range, style, opts, 
      function (error, data, response) {
        if (error) throw error;
    });
})();