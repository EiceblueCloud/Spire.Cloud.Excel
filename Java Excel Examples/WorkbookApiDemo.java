import java.io.File;
import spire.cloud.excel.sdk.api.WorkbookApi;
import spire.cloud.excel.sdk.model.*;

public class WorkbookApiDemo {
	   static String appId = "your id";
     static String appKey = "your key";
     
     public void GetWorkbook() throws Exception{
         WorkbookApi workbookApi = new WorkbookApi(appId, appKey);
         String name = "GetWorkbook.xlsx";
         String password = null;
         String storage = null;
         String folder = "input";
         Workbook  response = workbookApi.getWorkbook(name, password, storage, folder);       
     }
     public void PostWorkbook() throws Exception{
    	 WorkbookApi workbookApi = new WorkbookApi(appId, appKey);
    	 String name = "PostWorkbook.xlsx";
    	 String inputFilePath = "D:/input/PostWorkbook.xlsx";
    	 File data = new File(inputFilePath);
         String inputPassword = null;
         String password = null;
         String storage = null;
         String folder = "output";
         Workbook response = workbookApi.postWorkbook(name, data, inputPassword, password, storage, folder);
     }
     public void PostWorkbookTestPassword() throws Exception{
    	 WorkbookApi workbookApi = new WorkbookApi(appId, appKey);
         String name = "PostWorkbookTestPassword.xlsx";
         String inputFilePath = "D:/input/postWorkbookConvert.xlsx";
         File data = new File(inputFilePath);
         String inputPassword = "123";
         String password = null;
         String storage = null;
         String folder = "output";
         Workbook response = workbookApi.postWorkbook(name, data, inputPassword, password, storage, folder);  
     }
     public void PostWorkbookFromSource() throws Exception{
    	 WorkbookApi workbookApi = new WorkbookApi(appId, appKey);
         String name = "PostWorkbookFromSource.xlsx";
         String sourcePath = "input";
         String sourcePassword = "123";
         String sourceStorage = null;
         String password = "test";
         String storage = null;
         String folder = "output";
         Workbook response = workbookApi.postWorkbookFromSource(name, sourcePath, sourcePassword, sourceStorage, password, storage, folder);
     }
     public void PostWorkbookSaveAsXls() throws Exception{
    	 WorkbookApi workbookApi = new WorkbookApi(appId, appKey);
         String name = "PostWorkbookSaveAsXls.xlsx";
         String format = ExportFormat.XLS.toString();
         ExportOptions options = null;
         String password = null;
         String storage = null;
         String folder = "input";
         File response = workbookApi.postWorkbookSaveAs(name, format, options, password, storage, folder);
     }
     public void PutWorkbookSaveAsXps() throws Exception{
    	 WorkbookApi workbookApi = new WorkbookApi(appId, appKey);
    	 String name = "PostWorkbookSaveAsXps.xlsx";
    	 String outputFilePath = "output/PutWorkbookSaveAsXps_output.xps";
    	 String format = ExportFormat.XPS.toString();
         ExportOptions options = null;
         String password = null;
         String storage = null;
         String inputfolder = "input";
         workbookApi.putWorkbookSaveAs(name, outputFilePath, format, options, password, storage, inputfolder);
     }
     public void PutWorkbookSaveAsTestPassword() throws Exception{
    	 WorkbookApi workbookApi = new WorkbookApi(appId, appKey);
    	 String name = "PutWorkbookSaveAsTestPassword.xlsx";
    	 String outputFilePath = "output/PutWorkbookSaveAsTestPassword_output.xps";
    	 String format = ExportFormat.XPS.toString();
         ExportOptions options = null;
         String password = "123";
         String storage = null;
         String inputfolder = "input";
         workbookApi.putWorkbookSaveAs(name, outputFilePath, format, options, password, storage,inputfolder);
     }
}
