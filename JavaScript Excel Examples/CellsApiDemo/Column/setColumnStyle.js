(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.CellsApi(apiClient);

    var name= "test1.xlsx";
    var sheetName = "Sheet1";
    var columnIndex = 2;
    var opts = { folder: "input" };
    var style = new SpirecloudExcel.Style();
    style.Font = new SpirecloudExcel.Font();
    style.Font.IsBold = true;
    style.Font.IsItalic = true;
    style.Font.Underline = "Single";
    style.Font.Name = "Algerian";
    style.Font.Size = 8;
    style.IsTextWrapped = true;
    style.Pattern = "Solid";
    style.HorizontalAlignment = "Center";
    style.BackgroundColor = new SpirecloudExcel.Color();
    style.BackgroundColor.A = 255;
    style.BackgroundColor.R = 0;
    style.BackgroundColor.G = 255;
    style.BackgroundColor.B = 0;

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
    style.BorderCollection[1].Color.A = 105;
    style.BorderCollection[1].Color.R = 0;
    style.BorderCollection[1].Color.G = 100;
    style.BorderCollection[1].Color.B = 0;
    style.BorderCollection[1].BorderType = "EdgeRight";
    instance.setColumnStyle(name, sheetName, columnIndex, style, opts, 
        function (error, data, response) {
            if (error) throw error;
        });
})();