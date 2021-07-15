import spire.cloud.excel.sdk.*;
import spire.cloud.excel.sdk.api.WorkbookApi;
import spire.cloud.excel.sdk.model.*;

import java.io.File;

public class WorkbookApiDemo {
    static String appId = "your id";
    static String appKey = "your key";
    static String baseUrl = "https://api.e-iceblue.cn";
    static Configuration configuration = new Configuration(appId, appKey, baseUrl);
    static WorkbookApi workbookApi = new WorkbookApi(configuration);

    public static void convertWorkbook() throws ApiException {
        String name = "ConvertWorkbook.xlsx";
        String format = ExportFormat.XLSX.toString();//Supported formats: XLSX/XLS/XLSB/ODS/PDF/XPS/PS/PCL
        ExportOptions options = null;
        String password = null;
        String storage = null;
        String folder = "input";
        File response = workbookApi.convertWorkbook(name, format, options, password, storage, folder);
    }

    public static void convertWorkbookToPath() throws ApiException {
        String name = "ConvertWorkbookToPath.xlsx";
        String outPath = "output/ConvertWorkbookToPath.xlsx";
        String format = ExportFormat.XLSX.toString();//Supported formats: XLSX/XLS/XLSB/ODS/PDF/XPS/PS/PCL
        ExportOptions options = null;
        String password = null;
        String storage = null;
        String folder = "input";
        workbookApi.convertWorkbookToPath(name, outPath, format, options, password, storage, folder);
    }

    public static void createWorkbook() throws ApiException {
        String name = "CreateWorkbook.xlsx";
        String filePath = "data/Sample.xlsx";
        File data=new File(filePath);
        String inputPassword = null;
        String password = null;
        String storage = null;
        String folder = "output";
        Workbook response = workbookApi.createWorkbook(name, data, inputPassword, password, storage, folder);
    }

    public static void createWorkbookFromSource() throws ApiException {
        String name = "FromSource.xlsx";
        String sourcePath = "input";
        String sourcePassword = null;
        String sourceStorage = null;
        String password = null;
        String storage = null;
        String folder = "output";
        Workbook response = workbookApi.createWorkbookFromSource(name, sourcePath, sourcePassword, sourceStorage, password, storage, folder);
    }

    public static void decryptDocument() throws ApiException {
        String storage = null;
        String folder = "input";
        String name = "DecryptDocument.xlsx";
        WorkbookEncryptionRequest encryption = new WorkbookEncryptionRequest();
        encryption.password("123");
        workbookApi.decryptDocument(name, encryption, folder, storage);
    }

    public static void encryptDocument() throws ApiException {
        String name = "EncryptDocument.xlsx";
        String storage = null;
        String folder = "input";
        WorkbookEncryptionRequest encryption = new WorkbookEncryptionRequest();
        encryption.password("123");
        workbookApi.encryptDocument(name, encryption, folder, storage);
    }

    public static void getWorkbook() throws ApiException {
        String name = "GetWorkbook.xlsx";
        String password = null;
        String storage = null;
        String folder = "input";
        Workbook response = workbookApi.getWorkbook(name, password, storage, folder);
    }

    public static void protectDocument() throws ApiException {
        String name = "ProtectDocument.xlsx";
        String folder = "input";
        String storage = null;
        WorkbookProtectionRequest protection = new WorkbookProtectionRequest();
        protection.password("123");
        protection.protectionType("ALL");//Available value: ALL, READONLY, STRUCTURE, WINDOWS
        workbookApi.protectDocument(name, protection, folder, storage);
    }

    public static void unProtectDocument() throws ApiException {
        String name = "UnProtectDocument.xlsx";
        String folder = "input";
        String storage = null;
        WorkbookProtectionRequest protection = new WorkbookProtectionRequest();
        protection.password("123");
        protection.protectionType("READONLY");//Available value: ALL, READONLY, STRUCTURE, WINDOWS
        workbookApi.unProtectDocument(name, protection, folder, storage);
    }
}
