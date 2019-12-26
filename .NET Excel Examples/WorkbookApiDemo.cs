using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Cloud.Excel.Sdk.Api;
using Spire.Cloud.Excel.Sdk.Client;
using Spire.Cloud.Excel.Sdk.Model;

namespace CloudExcel
{
    class WorkbookApiDemo
    {
       static String appId = "your id";
       static String appKey = "your key";
        public static void GetWorkbook()
        {
            Configuration configuration = new Configuration("https://api.e-iceblue.cn", appId, appKey);
            WorkbookApi workbookApi = new WorkbookApi(configuration);
            string name = "GetWorkbook.xlsx";
            string password = null;
            string storage = null;
            string folder = "input";
            var response = workbookApi.GetWorkbook(name, password, storage, folder);       
        }
        public static void PostWorkbook()
        {
            Configuration configuration = new Configuration("https://api.e-iceblue.cn", appId, appKey);
            WorkbookApi workbookApi = new WorkbookApi(configuration);
            string name = "PostWorkbook.xlsx";
            string inputFilePath = "D:/input/PostWorkbook.xlsx";
            System.IO.Stream data = new FileStream(inputFilePath, FileMode.Open);
            string inputPassword = null;
            string password = null;
            string storage = null;
            string folder = "output";
            var response = workbookApi.PostWorkbook(name, data, inputPassword, password, storage, folder);
            data.Close();
        }
        public static void PostWorkbookTestPassword()
        {
            Configuration configuration = new Configuration("https://api.e-iceblue.cn", appId, appKey);
            WorkbookApi workbookApi = new WorkbookApi(configuration);
            string name = "PostWorkbookTestPassword.xlsx";
            string inputFilePath = "D:/input/postWorkbookConvert.xlsx";
            System.IO.Stream data = new FileStream(inputFilePath, FileMode.Open);
            string inputPassword = "123";
            string password = null;
            string storage = null;
            string folder = "output";
            var response = workbookApi.PostWorkbook(name, data, inputPassword, password, storage, folder);  
            data.Close();
        }
        public static void PostWorkbookFromSource()
        {
            Configuration configuration = new Configuration("https://api.e-iceblue.cn", appId, appKey);
            WorkbookApi workbookApi = new WorkbookApi(configuration);
            string name = "PostWorkbookFromSource.xlsx";
            string sourcePath = "input";
            string sourcePassword = "123";
            string sourceStorage = null;
            string password = "test";
            string storage = null;
            string folder = "output";
            var response = workbookApi.PostWorkbookFromSource(name, sourcePath, sourcePassword, sourceStorage, password, storage, folder);
        }
        public void PostWorkbookSaveAsXls()
        {
            Configuration configuration = new Configuration("https://api.e-iceblue.cn", appId, appKey);
            WorkbookApi workbookApi = new WorkbookApi(configuration);
            string name = "PostWorkbookSaveAsXls.xlsx";
            string format = ExportFormat.Xls.ToString();
            ExportOptions options = null;
            string password = null;
            string storage = null;
            string folder = "input";
            var response = workbookApi.PostWorkbookSaveAs(name, format, options, password, storage, folder);
        }
        public static void PutWorkbookSaveAsXps()
        {
            Configuration configuration = new Configuration("https://api.e-iceblue.cn", appId, appKey);
            WorkbookApi workbookApi = new WorkbookApi(configuration);
            string name = "PostWorkbookSaveAsXps.xlsx";
            string outputFilePath = @"output/PutWorkbookSaveAsXps_output.xps";
            string format = ExportFormat.Xps.ToString();
            ExportOptions options = null;
            string password = null;
            string storage = null;
            string inputfolder = "input";
            workbookApi.PutWorkbookSaveAs(name, outputFilePath, format, options, password, storage, inputfolder);
        }
        public static void PutWorkbookSaveAsTestPassword()
        {
            Configuration configuration = new Configuration("https://api.e-iceblue.cn", appId, appKey);
            WorkbookApi workbookApi = new WorkbookApi(configuration);
            string name = "PutWorkbookSaveAsTestPassword.xlsx";
            string outputFilePath = @"output/PutWorkbookSaveAsTestPassword_output.xps";
            string format = ExportFormat.Xps.ToString();
            ExportOptions options = null;
            string password = "123";
            string storage = null;
            string inputfolder = "input";
            workbookApi.PutWorkbookSaveAs(name, outputFilePath, format, options, password, storage,inputfolder);
        }
    }
}
