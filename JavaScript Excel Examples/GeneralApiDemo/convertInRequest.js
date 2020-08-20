(function (){
          var appId = "your id";
          var appKey = "your key";
          var baseUrl = "https://api.e-iceblue.cn";
      
          var SpirecloudExcel = require('../../src/index');
          var apiClient = new SpirecloudExcel.ApiClient(appId, appKey,baseUrl);
          var instance = new SpirecloudExcel.GeneralApi(apiClient);
          
          //upload the file from local disk
          var fs = require('fs');
          var filename = fs.createReadStream("F:/convertInRequest.xlsx");
          //supports formats: Xlsx/Xls/Xlsb/Ods/Pdf/Xps/Ps/Pcl
          var format = "Xlsx";
          var opts = {file: filename, password:""};
          instance.convertInRequest(format, opts, function (error, data, response) {
              if (error) throw error;
          });
})();