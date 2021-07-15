import spire.cloud.excel.sdk.*;
import spire.cloud.excel.sdk.api.RangesApi;
import spire.cloud.excel.sdk.model.*;

import java.util.ArrayList;

public class RangesApiDemo {
    static String appId = "your id";
    static String appKey = "your key";
    static String baseUrl = "https://api.e-iceblue.cn";
    static Configuration configuration = new Configuration(appId, appKey, baseUrl);
    static RangesApi rangesApi = new RangesApi(configuration);

    public static void copyRanges() throws ApiException {
        String name = "CopyRanges.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        RangeCopyRequest rangeOperate = new RangeCopyRequest();
        rangeOperate.setOperate("copystyle");
        Range range = new Range();
        range.firstColumn(1);
        range.firstRow(1);
        range.rowCount(10);
        range.columnCount(10);
        Range range2 = new Range();
        range2.firstColumn(30);
        range2.firstRow(1);
        range2.rowCount(5);
        range2.columnCount(8);
        rangeOperate.source(range);
        rangeOperate.target(range2);
        rangesApi.copyRanges(name, sheetName, rangeOperate, folder, storage);
    }

    public static void getRangeValue() throws ApiException {
        String name = "GetRangeValue.xlsx";
        String sheetName = "Sheet1";
        String folder = "input";
        String storage = null;
        String nameRange = "C1:D3";
        RangeValue response = rangesApi.getRangeValue(name, sheetName, nameRange, null, null, null, null, folder, storage);
    }

    public static void mergeRange() throws ApiException {
        String name = "MergeRange.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        Range range = new Range();
        range.firstColumn(3);
        range.firstRow(3);
        range.rowCount(5);
        range.columnCount(8);
        rangesApi.mergeRange(name, sheetName, range, folder, storage);
    }

    public static void moveRange() throws ApiException {
        String name = "MoveRange.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        Range range = new Range();
        range.firstColumn(1);
        range.firstRow(1);
        range.rowCount(1);
        range.columnCount(1);
        int destRow = 25;
        int destColumn = 3;
        rangesApi.moveRange(name, sheetName, range, destRow, destColumn, folder, storage);
    }
    public static void setRangeColumnWidth() throws ApiException {
        String name = "SetRangeColumnWidth.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        Range range = new Range();
        range.firstColumn(1);
        range.firstRow(1);
        range.rowCount(2);
        range.columnCount(5);
        double value = 15;
        rangesApi.setRangeColumnWidth(name, sheetName, range, value, folder, storage);
    }
    public static void setRangeOutlineBorder() throws ApiException {
        String name = "SetRangeOutlineBorder.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        RangeSetOutlineBorderRequest rangeOperate =
                new RangeSetOutlineBorderRequest();
        rangeOperate.borderEdge("outline");//Available value: vertical, horizontal, outline, edgeleft and so on.
        rangeOperate.borderStyle("Medium");//Available value: Medium, DashDot, Thin, Dotted, Hair, MediumDashDot, MediumDashDotDot and so on.
        rangeOperate.borderColor(new Color(255, 255, 0, 0));

        Range range = new Range();
        range.firstColumn(1);
        range.firstRow(1);
        range.columnCount(10);
        range.rowCount(5);
        rangeOperate.range(range);
        rangesApi.setRangeOutlineBorder(name, sheetName, rangeOperate, folder, storage);
    }

    public static void setRangeRowHeight() throws ApiException {
        String name = "SetRangeRowHeight.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        Range range = new Range();
        range.firstColumn(1);
        range.firstRow(1);
        range.columnCount(5);
        range.rowCount(10);
        double value = 20;
        rangesApi.setRangeRowHeight(name, sheetName, range, value, folder, storage);
    }

    public static void setRangeStyle() throws ApiException {
        String name = "SetRangeStyle_1.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        RangeSetStyleRequest rangeOperate = new RangeSetStyleRequest();
        rangeOperate.setStyle(new Style());
        rangeOperate.getStyle().font(new Font());
        rangeOperate.getStyle().getFont().color(new Color(255, 255, 0, 0));
        rangeOperate.getStyle().getFont().underline( "single");
        rangeOperate.getStyle().getFont().size(12);
        rangeOperate.getStyle().getFont().name("楷体");
        rangeOperate.getStyle().getFont().isItalic(true);
        rangeOperate.getStyle().getFont().isBold(true);

        ArrayList<Border> lists = new ArrayList<Border>();
        Border border = new Border();
        border.setLineStyle("Medium");
        border.setBorderType("EdgeTop");
        border.setColor(new Color(255, 255, 0, 0));
        lists.add(border);

        Border border1 = new Border();
        border1.setLineStyle("DashDot");
        border1.setBorderType("EdgeRight");
        border1.setColor(new Color(255, 0, 255, 0));
        lists.add(border1);

        Border border2 = new Border();
        border2.setLineStyle("Thin");
        border2.setBorderType("EdgeLeft");
        border2.setColor(new Color(0, 160, 32, 240));
        lists.add(border2);
        rangeOperate.getStyle().setBorderCollection(lists);
        rangeOperate.getStyle().horizontalAlignment("left");//需要设置HorizontalAlignment该属性方可生效
        rangeOperate.getStyle().indentLevel(1);
        rangeOperate.getStyle().textDirection("RightToLeft");
        rangeOperate.getStyle().backgroundColor(new Color(0, 255, 255, 0));
        rangeOperate.getStyle().isTextWrapped(true);

        Range range = new Range();
        range.firstColumn(2);
        range.firstRow(1);
        range.columnCount(1);
        range.rowCount(10);
        rangeOperate.range(range);
        rangesApi.setRangeStyle(name, sheetName, rangeOperate, folder, storage);
    }

    public static void setRangeValue() throws ApiException {
        String name = "SetRangeValue.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet2";
        Range range = new Range();
        range.firstRow(1);
        range.firstColumn(1);
        range.columnCount(5);
        range.rowCount(2);
        String value = "E-iceblue";
        boolean isConverted = true;
        boolean setStyle = false;
        rangesApi.setRangeValue(name, sheetName, range, value, isConverted, setStyle, folder, storage);
    }

    public static void unmergeRange() throws ApiException {
        String name = "UnMergeRange.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        Range range = new Range();
        range.firstColumn(3);
        range.firstRow(3);
        range.rowCount(5);
        range.columnCount(8);
        rangesApi.unmergeRange(name, sheetName, range, folder, storage);
    }
}
