(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);
    
    var name = "setBackground.xlsx";
    var opts = {folder: "input",picPath: "logo.png"};
    var sheetName = "Sheet1";
    instance.setBackground(name, sheetName, opts, 
      function (error, data, response) {
        if (error) throw error;
    });
})();