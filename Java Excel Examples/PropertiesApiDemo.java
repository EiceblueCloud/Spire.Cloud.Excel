import spire.cloud.excel.sdk.Configuration;
import spire.cloud.excel.sdk.api.PropertiesApi;
import spire.cloud.excel.sdk.model.*;

import java.util.ArrayList;

public class PropertiesApiDemo {
    static String appId = "your id";
    static String appKey = "your key";
    static String baseUrl = "https://api.e-iceblue.cn";
    static Configuration configuration = new Configuration(appId, appKey, baseUrl);
    static PropertiesApi propertiesApi = new PropertiesApi(configuration);
    public static void deleteDocumentProperties() throws Exception{
        String name = "DeleteDocumentProperties.xlsx";
        String password = null;
        String folder = "input";
        String storage = null;

        DocumentProperties response = propertiesApi.deleteDocumentProperties(name, password, folder, storage);
    }

    public static void deleteDocumentProperty() throws Exception{
        String name = "DeleteDocumentProperty.xlsx";
        String password = null;
        String folder = "input";
        String storage = null;
        String propertyName = "Title";
        DocumentProperties response = propertiesApi.deleteDocumentProperty(name, propertyName, password, folder, storage);
    }

    public static void getDocumentProperties() throws Exception{
        String name = "GetDocumentProperties.xlsx";
        String password = null;
        String folder = "input";
        String storage = null;
        DocumentProperties response = propertiesApi.getDocumentProperties(name, password, folder, storage);
    }

    public static void getDocumentProperty() throws Exception{
        String name = "GetDocumentProperty.xlsx";
        String propertyName = "Title";
        String password = null;
        String folder = "input";
        String storage = null;
        DocumentProperty response = propertiesApi.getDocumentProperty(name, propertyName, password, folder, storage);
    }

    public static void setDocumentProperties() throws Exception{
        String name = "SetDocumentProperties.xlsx";
        String password = null;
        String folder = "input";
        String storage = null;
        DocumentProperties properties = new DocumentProperties();
        DocumentProperty property = new DocumentProperty("Keywords", "Set document properties.", false);
        DocumentProperty property2 = new DocumentProperty("Author", "eiceblue", false);
        DocumentProperty property3 = new DocumentProperty("Company", "冰蓝", true);
        DocumentProperty property4 = new DocumentProperty("Last saved by", "123456@iCloud.com", true);
        DocumentProperty property5 = new DocumentProperty("Status", "true", false);
        DocumentProperty property6 = new DocumentProperty("Company", "e-iceblue", true);
        DocumentProperty property7 = new DocumentProperty("作者", "冰蓝", false);
        properties.setList(new ArrayList<DocumentProperty>());
        properties.getList().add(property);
        properties.getList().add(property2);
        properties.getList().add(property3);
        properties.getList().add(property4);
        properties.getList().add(property5);
        properties.getList().add(property6);
        properties.getList().add(property7);
        DocumentProperties response = propertiesApi.setDocumentProperties(name, properties, password, folder, storage);
    }

    public static void setDocumentProperty() throws Exception{
        String name = "SetDocumentProperty.xlsx";
        String password = null;
        String folder = "input";
        String storage = null;
        String propertyName = "Keywords";
        DocumentProperty property = new DocumentProperty("Keywords", "SetDocumentProperty_xlsx", true);
        DocumentProperty response = propertiesApi.setDocumentProperty(name, propertyName, property, password, folder, storage);
    }
}
