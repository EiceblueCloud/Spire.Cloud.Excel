(function (){
    var appId = "your id";
    var appKey = "your key"; 
    var baseUrl = "https://api.e-iceblue.cn"; 

    var SpirecloudExcel= require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorkbookApi(apiClient);

    var name = "CreateWorkbookEmpty.xlsx";
    var opts = {folder: "workbook"};
    instance.createWorkbook(name, opts,
        function (error, data, response) {
            if (error) throw error;
        });
})();