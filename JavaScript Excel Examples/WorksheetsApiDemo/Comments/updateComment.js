(function (){
    var appId = "your id";
    var appKey = "your key";
    var baseUrl = "https://api.e-iceblue.cn";

    var SpirecloudExcel = require('../../../src/index');
    var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
    var instance = new SpirecloudExcel.WorksheetsApi(apiClient);
    
    var name = "updateComment.xlsx";
    var opts = { folder: "input" };
    var sheetName = "Sheet1";
    var cellName = "F9";
    var comment = new SpirecloudExcel.Comment();
    comment.Author = "jerry";
    comment.Note = "Hello,Comments";
    comment.TextOrientationType = "TopToBottom";
    comment.IsVisible = false;
    comment.AutoSize = true;
    instance.updateComment(name, sheetName, cellName, comment, opts,
            function (error, data, response) {
        if (error) throw error;
    });
})();