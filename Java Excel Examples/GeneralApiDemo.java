import spire.cloud.excel.sdk.Configuration;
import spire.cloud.excel.sdk.api.GeneralApi;
import spire.cloud.excel.sdk.model.ExportFormat;

import java.io.File;

public class GeneralApiDemo {
    static String appId = "your id";
    static String appKey = "your key";
    static String baseUrl = "https://api.e-iceblue.cn";
    static Configuration configuration = new Configuration(appId, appKey, baseUrl);
    static GeneralApi generalApi = new GeneralApi(configuration);
    public static void convertInRequest() throws Exception {
        //Supported formats: Xlsx/Xls/Xlsb/Ods/Pdf/Xps/Ps/Pcl
        String format = ExportFormat.XLSX.toString();
        File document = new File("D:/inputFile/charts.xlsx");
        String password = null;
        File response = generalApi.convertInRequest(format, document, password);
    }

    public static void convertInRequestToPath() throws Exception{
        //Supported formats: Xlsx/Xls/Xlsb/Ods/Pdf/Xps/Ps/Pcl
        String format = ExportFormat.XLSX.toString();
        String outPath = "output/ConvertInRequestToPath.xlsx";
        File document = new File("D:/inputFile/charts.xlsx");
        String password = null;
        generalApi.convertInRequestToPath(format, outPath, document, password);
    }
}
