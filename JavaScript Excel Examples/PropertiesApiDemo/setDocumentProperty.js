(function (){
    var appId = "your id";
    var appKey = "your key"; 
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel= require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.PropertiesApi(apiClient);

    var name = "Workbook.xlsx";
    var propertyName = "Keywords";
    var property = new SpirecloudExcel.DocumentProperty();
    property.Name = "Keywords";
    property.Value = "SetDocumentProperty_xlsx"
    property.BuiltIn = true;

    var opts = {folder : "input", property: property};
    instance.setDocumentProperty(name, propertyName, opts,
        function (error, data, response) {
            if (error) throw error;
        });
})();