(function (){
    var appId = "your id";
    var appKey = "your key";  
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel= require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorkbookApi(apiClient);

    var name = "Workbook.xlsx";
    var opts = {folder : "input"};
    var encryption = new SpirecloudExcel.WorkbookEncryptionRequest();
    encryption.password = "123";
    instance.encryptDocument(name, encryption, opts,
        function (error, data, response) {
            if (error) throw error;
        });
})();