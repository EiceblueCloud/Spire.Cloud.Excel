using Spire.Cloud.Excel.Sdk.Api;
using Spire.Cloud.Excel.Sdk.Client;
using Spire.Cloud.Excel.Sdk.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CloudExcel
{
    class PropertiesApiDemo
    {
        static String appId = "your id";
        static String appKey = "your key";
        static String baseUrl = "https://api.e-iceblue.cn";
        static Configuration configuration = new Configuration(appId, appKey, baseUrl);
        static PropertiesApi propertiesApi = new PropertiesApi(configuration);

        public static void DeleteDocumentProperties()
        {
            string name = "DeleteDocumentProperties.xlsx";
            string password = null;
            string folder = "input";
            string storage = null;

            var response = propertiesApi.DeleteDocumentProperties(name, password, folder, storage);
        }

        public static void DeleteDocumentProperty()
        {
            string name = "DeleteDocumentProperty.xlsx";
            string password = null;
            string folder = "input";
            string storage = null;
            string propertyName = "Title";
            var response = propertiesApi.DeleteDocumentProperty(name, propertyName, password, folder, storage);
        }

        public static void GetDocumentProperties()
        {
            string name = "GetDocumentProperties.xlsx";
            string password = null;
            string folder = "input";
            string storage = null;
            var response = propertiesApi.GetDocumentProperties(name, password, folder, storage);
        }

        public static void GetDocumentProperty()
        {
            string name = "GetDocumentProperty.xlsx";
            string propertyName = "Title";
            string password = null;
            string folder = "input";
            string storage = null;
            var response = propertiesApi.GetDocumentProperty(name, propertyName, password, folder, storage);
        }

        public static void SetDocumentProperties()
        {
            string name = "SetDocumentProperties.xlsx";
            string password = null;
            string folder = "input";
            string storage = null;
            DocumentProperties properties = new DocumentProperties();
            DocumentProperty property = new DocumentProperty("Keywords", "Set document properties.", false);
            DocumentProperty property2 = new DocumentProperty("Author", "eiceblue", false);
            DocumentProperty property3 = new DocumentProperty("Company", "冰蓝", true);
            DocumentProperty property4 = new DocumentProperty("Last saved by", "123456@iCloud.com", true);
            DocumentProperty property5 = new DocumentProperty("Status", "true", false);
            DocumentProperty property6 = new DocumentProperty("Company", "e-iceblue", true);
            DocumentProperty property7 = new DocumentProperty("作者", "冰蓝", false);
            properties.List = new List<DocumentProperty>();
            properties.List.Add(property);
            properties.List.Add(property2);
            properties.List.Add(property3);
            properties.List.Add(property4);
            properties.List.Add(property5);
            properties.List.Add(property6);
            properties.List.Add(property7);
            var response = propertiesApi.SetDocumentProperties(name, properties, password, folder, storage);
        }

        public static void SetDocumentProperty()
        {
            string name = "SetDocumentProperty.xlsx";
            string password = null;
            string folder = "input";
            string storage = null;
            string propertyName = "Keywords";
            DocumentProperty property = new DocumentProperty("Keywords", "SetDocumentProperty_xlsx", true);
            var response = propertiesApi.SetDocumentProperty(name, propertyName, property, password, folder, storage);
        }
    }
}
