(function (){
    var appId = "your id";
    var appKey = "your key"; 
    
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);

    var name = "UpdateWorksheetProperty.xlsx";
    var opts = {
        folder: "input"
    };
    var sheetName = "Sheet2";
    var sheet = new SpirecloudExcel.Worksheet();
    sheet.Zoom = 50;
    sheet.Index = 3;
    sheet.IsGridlinesVisible = false; 
    sheet.IsSelected = true;
    sheet.IsProtected = true;
    sheet.TabColor = new SpirecloudExcel.Color(); 
    sheet.TabColor.A = 255;
    sheet.TabColor.R = 255;
    sheet.TabColor.G = 0;
    sheet.TabColor.B = 0;
    instance.updateWorksheetProperty(name, sheetName, sheet, opts, function (error, data, response) {
            if (error) throw error;
        });
})();