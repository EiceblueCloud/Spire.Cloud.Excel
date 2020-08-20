(function (){
    var appId = "your id";
    var appKey = "your key"; 
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel= require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.PropertiesApi(apiClient);

    var name = "DocumentProperties.xlsx";
    var propertyName = "Title";
    var opts = {folder : "input"};
    instance.getDocumentProperty(name, propertyName, opts,
        function (error, data, response) {
            if (error) throw error;
        });
})();