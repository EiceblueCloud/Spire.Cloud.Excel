using Spire.Cloud.Excel.Sdk.Api;
using Spire.Cloud.Excel.Sdk.Client;
using Spire.Cloud.Excel.Sdk.Model;
using System;
using System.IO;

namespace CloudExcel
{
    class WorkbookApiDemo
    {
        static String appId = "your id";
        static String appKey = "your key";
        static String baseUrl = "https://api.e-iceblue.cn";
        static Configuration configuration = new Configuration(appId, appKey, baseUrl);
        static WorkbookApi workbookApi = new WorkbookApi(configuration);
        public static void ConvertWorkbook()
        {
            //Supported formats: Xlsx/Xls/Xlsb/Ods/Pdf/Xps/Ps/Pcl
            string name = "ConvertWorkbook.xlsx";
            string format = ExportFormat.Xlsx.ToString();
            ExportOptions options = null;
            string password = null;
            string storage = null;
            string folder = "input";
            var response = workbookApi.ConvertWorkbook(name, format, options, password, storage, folder);
        }

        public static void ConvertWorkbookToPath()
        {
            string name = "ConvertWorkbookToPath.xlsx";
            string outPath = "output/ConvertWorkbookToPath.xlsx";
            //Supported formats: Xlsx/Xls/Xlsb/Ods/Pdf/Xps/Ps/Pcl
            string format = ExportFormat.Xlsx.ToString();
            ExportOptions options = null;
            string password = null;
            string storage = null;
            string folder = "input";
            workbookApi.ConvertWorkbookToPath(name, outPath, format, options, password, storage, folder);
        }

        public static void CreateWorkbook()
        {
            string name = "CreateWorkbook.xlsx";
            Stream data = new FileStream("inputFile/charts.xlsx", FileMode.Open);
            string inputPassword = null;
            string password = null;
            string storage = null;
            string folder = "output";
            var response = workbookApi.CreateWorkbook(name, data, inputPassword, password, storage, folder);
        }

        public static void CreateWorkbookFromSource()
        {
            string name = "FromSource.xlsx";
            string sourcePath = "input";
            string sourcePassword = null;
            string sourceStorage = null;
            string password = null;
            string storage = null;
            string folder = "output";
            var response = workbookApi.CreateWorkbookFromSource(name, sourcePath, sourcePassword, sourceStorage, password, storage, folder);
        }

        public static void DecryptDocument()
        { 
            string storage = null;
            string folder = "input";
            string name = "DecryptDocument.xlsx";
            WorkbookEncryptionRequest encryption = new WorkbookEncryptionRequest();
            encryption.Password = "123";

            workbookApi.DecryptDocument(name, encryption, folder, storage);
        }

        public static void EncryptDocument()
        {
            string name = "EncryptDocument.xlsx";
            string storage = null;
            string folder = "input";
            WorkbookEncryptionRequest encryption = new WorkbookEncryptionRequest();
            encryption.Password = "123";
            workbookApi.EncryptDocument(name, encryption, folder, storage);
        }

        public static void GetWorkbook()
        {
            string name = "GetWorkbook.xlsx";//Supported formats: xlsx/xls/xlsb/xlsm/xltm/xltx/ods
            string password = null;
            string storage = null;
            string folder = "input";
            var response = workbookApi.GetWorkbook(name, password, storage, folder);
        }

        public static void ProtectDocument()
        {
            string name = "ProtectDocument.xlsx";
            string folder = "input";
            string storage = null;
            WorkbookProtectionRequest protection = new WorkbookProtectionRequest();
            protection.Password = "123";
            protection.ProtectionType = "ALL";// ALL, READONLY, STRUCTURE, WINDOWS
            workbookApi.ProtectDocument(name, protection, folder, storage);
        }

        public static void UnProtectDocument()
        {
            string name = "UnProtectDocument.xlsx";
            string folder = "input";
            string storage = null;
            WorkbookProtectionRequest protection = new WorkbookProtectionRequest();
            protection.Password = "123";
            protection.ProtectionType = "READONLY";// ALL, READONLY, STRUCTURE, WINDOWS
            workbookApi.UnProtectDocument(name, protection, folder, storage);
        }
    }
}
