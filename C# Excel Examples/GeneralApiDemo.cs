using Spire.Cloud.Excel.Sdk.Api;
using Spire.Cloud.Excel.Sdk.Client;
using Spire.Cloud.Excel.Sdk.Model;
using System;
using System.IO;

namespace CloudExcel
{
    class GeneralApiDemo
    {
        static String appId = "your id";
        static String appKey = "your key";
        static String baseUrl = "https://api.e-iceblue.cn";
        static Configuration configuration = new Configuration(appId, appKey, baseUrl);
        static GeneralApi generalApi = new GeneralApi(configuration);
        public static void ConvertInRequest()
        {
            //Supported formats: Xlsx/Xls/Xlsb/Ods/Pdf/Xps/Ps/Pcl
            string format = ExportFormat.Xlsx.ToString();
            Stream document = new FileStream("inputFile/charts.xlsx", FileMode.Open);
            string password = null;
            var response = generalApi.ConvertInRequest(format, document, password);
        }

        public static void ConvertInRequestToPath()
        { 
            //Supported formats: Xlsx/Xls/Xlsb/Ods/Pdf/Xps/Ps/Pcl
            string format = ExportFormat.Xlsx.ToString();
            string outPath = "output/ConvertInRequestToPath.xlsx";
            Stream document = new FileStream("/inputFile/charts.xlsx", FileMode.Open);
            string password = null;
            generalApi.ConvertInRequestToPath(format, outPath, document, password);
            document.Close();
        }
    }
}
