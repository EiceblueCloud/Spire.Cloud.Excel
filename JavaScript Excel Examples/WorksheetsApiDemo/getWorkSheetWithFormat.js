(function (){
    var appId = "your id";
    var appKey = "your key"; 
    
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);

    var name = "GetWorkSheetWithFormat_html.xlsx";
    var sheetName = "Sheet2";
    var format = "html"; //Available values :html, pdf, png, jpeg, bmp, emf, tiff
    var opts = {
        folder: "input"
    };
    instance.getWorkSheetWithFormat(name, sheetName, format, opts, function (error, data, response) {
            if (error) throw error;
        });
})();