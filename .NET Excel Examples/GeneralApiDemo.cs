using Spire.Cloud.Excel.Sdk.Api;
using Spire.Cloud.Excel.Sdk.Client;
using Spire.Cloud.Excel.Sdk.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CloudExcel
{
    class GeneralApiDemo
    {
       static String appId = "your id";
       static String appKey = "your key";
        public static void PostWorkbookConvert()
        {
            Configuration configuration = new Configuration("https://api.e-iceblue.cn", appId, appKey);
            GeneralApi generalApi = new GeneralApi(configuration);
            string format = ExportFormat.Pdf.ToString();
            string inputFilePath = "D:/input/postWorkbookConvert.xlsx";
            System.IO.Stream document = new FileStream(inputFilePath, FileMode.Open);
            string password = null;

            var response = generalApi.PostWorkbookConvert(format, document, password);
            document.Close();
        }

        public static void PutWorkbookConvert()
        {
            Configuration wordConfiguration = new Configuration("https://api.e-iceblue.cn", appId, appKey);
            GeneralApi generalApi = new GeneralApi(wordConfiguration);
            string format = ExportFormat.Xps.ToString();
            string inputFilePath = "D:/input/putWorkbookConvert.xlsx";
            System.IO.Stream document = new FileStream(inputFilePath, FileMode.Open);
            string outputFilePath = "output/putWorkbookConvert_output.xps";
            string password = null;
            generalApi.PutWorkbookConvert(format,outputFilePath,document, password);
            document.Close();
        }
    }
}