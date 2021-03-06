(function (){
    var appId = "your id";
    var appKey = "your key"; 
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel= require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorkbookApi(apiClient);

    var name = "Workbook.xlsx";
    var opts = {folder : "input"};
    instance.getWorkbook(name, opts, 
        function (error, data, response) {
            if (error) throw error;
        });
})();