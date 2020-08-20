(function (){
    var appId = "your id";
    var appKey = "your key"; 
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel= require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorkbookApi(apiClient);

    var name = "Workbook.xlsx";
    var protection = new SpirecloudExcel.WorkbookProtectionRequest();
    protection.password = "123";
    protection.ProtectionType = "STRUCTURE";//Available types: ALL, READONLY, STRUCTURE, WINDOWS
    var opts = {folder : "input"};
    instance.protectDocument(name, protection, opts,
        function (error, data, response) {
            if (error) throw error;
        });
})();