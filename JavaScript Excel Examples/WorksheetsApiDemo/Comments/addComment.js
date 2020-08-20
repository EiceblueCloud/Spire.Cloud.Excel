(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);
    
    var name = "addComment.xlsx";
    var opts = {folder: "input"};
    var sheetName = "Sheet1";
    var cellName = "C2:D4";
    var comment = new SpirecloudExcel.Comment();
    comment.author = "jerry";
    comment.note = "hello";
    comment.autoSize = true;
    instance.addComment(name, sheetName, cellName, comment, opts,
        function (error, data, response) {
        if (error) throw error;
    });
})();