(function (){
    var appId = "your id";
    var appKey = "your key"; 
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel= require('../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorkbookApi(apiClient);

    var name = "Workbook.xlsx";
    var format = "Xlsx"; //Available formats: Xlsx, Xls, Xlsb, Ods, Pdf, Xps, Ps, Pcl
    var opts = {folder : "input", options :new SpirecloudExcel.ExportOptions()};
    instance.convertWorkbook(name, format, opts,
        function (error, data, response) {
            if (error) throw error;
        });
})();