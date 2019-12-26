import java.io.File;
import spire.cloud.excel.sdk.api.GeneralApi;
import spire.cloud.excel.sdk.model.ExportFormat;

public class GeneralApiDemo {
	  static String appId = "your id";
    static String appKey = "your key";
    public void PostWorkbookConvert() throws Exception {
        GeneralApi generalApi = new GeneralApi(appId, appKey);
        String format = ExportFormat.PDF.toString();
        String inputFilePath = "D:/input/postWorkbookConvert.xlsx";
        File data = new File(inputFilePath);
        String password = null;
        File response = generalApi.postWorkbookConvert(format, data, password);
    }

    public void PutWorkbookConvert() throws Exception{
    	GeneralApi generalApi = new GeneralApi(appId, appKey);
        String format = ExportFormat.XPS.toString();
        String inputFilePath = "D:/input/putWorkbookConvert.xlsx";
        File data = new File(inputFilePath);
        String outputFilePath = "output/putWorkbookConvert_output.xps";
        String password = null;
        generalApi.putWorkbookConvert(format,outputFilePath,data, password);
    }
}
