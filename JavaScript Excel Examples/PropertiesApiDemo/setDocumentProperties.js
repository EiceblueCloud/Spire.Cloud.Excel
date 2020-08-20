(function (){
    var appId = "your id";
    var appKey = "your key"; 
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel= require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.PropertiesApi(apiClient);

    var name = "Workbook.xlsx";
    var properties = new SpirecloudExcel.DocumentProperties();
    properties.List = new Array();
    var property1 = new SpirecloudExcel.DocumentProperty();
    property1.Name = "Keywords";
    property1.Value = "Set document properties.";
    property1.BuiltIn = false;
    properties.List.push(property1);
    var property2 = new SpirecloudExcel.DocumentProperty();
    property2.Name = "Author";
    property2.Value = "eiceblue";
    property2.BuiltIn = false;
    properties.List.push(property2);
    var property3 = new SpirecloudExcel.DocumentProperty();
    property3.Name = "Company";
    property3.Value = "冰蓝";
    property3.BuiltIn = true;
    properties.List.push(property3);
    var property4 = new SpirecloudExcel.DocumentProperty();
    property4.Name = "LastSavedBy";
    property4.Value = "123456@iCloud.com.";
    property4.BuiltIn = true;
    properties.List.push(property4);
    var property5 = new SpirecloudExcel.DocumentProperty();
    property5.Name = "Status";
    property5.Value = "true";
    property5.BuiltIn = false;
    properties.List.push(property5);
    var property6 = new SpirecloudExcel.DocumentProperty();
    property6.Name = "Company";
    property6.Value = "e-iceblue";
    property6.BuiltIn = true;
    properties.List.push(property6);
    var property7 = new SpirecloudExcel.DocumentProperty();
    property7.Name = "作者";
    property7.Value = "e-冰蓝";
    property7.BuiltIn = false;
    properties.List.push(property7);

    var opts = {folder : "input",'properties': properties};
    instance.setDocumentProperties(name, opts, 
        function (error, data, response) {
            if (error) throw error;
        });
})();