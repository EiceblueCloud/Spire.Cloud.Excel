import spire.cloud.excel.sdk.Configuration;
import spire.cloud.excel.sdk.api.CellsApi;
import spire.cloud.excel.sdk.model.*;

import java.io.File;
import java.util.*;


public class CellsApiDemo {
    static String appId = "your id";
    static String appKey = "your key";
    static String baseUrl = "https://api.e-iceblue.cn";
    static Configuration configuration = new Configuration(appId, appKey, baseUrl);
    static CellsApi cellsApi = new CellsApi(configuration);
    public static void mergeCells() throws Exception{
        String name = "MergeCells.xlsx";
        String folder = "input";
        String storage = null;

        String sheetName = "Sheet4";
        int startRow = 2;
        int startColumn = 3;
        int totalRows = 4;
        int totalColumns = 4;
        cellsApi.mergeCells(name, sheetName, startRow, startColumn, totalRows, totalColumns, folder, storage);
    }

    public static void unMergeCells() throws Exception{
        String name = "UnMergeCells.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet2";
        int startRow = 3;
        int startColumn = 1;
        int totalRows = 1;
        int totalColumns = 2;
        cellsApi.unmergeCells(name, sheetName, startRow, startColumn, totalRows, totalColumns, folder, storage);
    }

    public static void getCellStyle() throws Exception{
        String name = "GetCellStyle.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet4";
        String cellName = "A4";
        Style response = cellsApi.getCellStyle(name, sheetName, cellName, folder, storage);
    }

    public static void setCellStyle() throws Exception{
        String name = "SetCellStyle.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet4";
        String cellName = "C8";
        Style style = new Style();
        Font font = new Font();
        font.setIsBold(true);
        font.setIsBold(true);
        font.setIsItalic(true);
        font.setUnderline("Single");
        font.setName("Algerian");
        font.setSize(8);
        style.setFont(font);
        style.setBackgroundColor(new Color(255,0,255,0));
        style.setHorizontalAlignment("Center");
        List<Border> list1= new ArrayList<Border>();
        Border border1 = new Border();
        border1.setLineStyle("Medium");
        border1.setColor(new Color(255, 255, 0, 0));
        border1.setBorderType("EdgeTop");
        list1.add(border1);
        Border border2 = new Border();
        border2.setLineStyle("DashDot");
        border2.setColor(new Color(255, 0, 255, 0));
        border2.setBorderType("EdgeRight");
        list1.add(border2);
        style.setBorderCollection(list1);
        Style response = cellsApi.setCellStyle(name, sheetName, cellName, style, folder, storage);
    }

    public static void calculateCell() throws Exception{
        String name = "CalculateCell.xlsx";
        String storage = null;
        String folder = "input";
        String sheetName = "Sheet5";
        String cellName = "G6";
        CalculationOptions options = new CalculationOptions();
        options.setCalcStackSize(0);
        options.setPrecisionStrategy("");
        options.setIgnoreError(true);
        options.setRecursive(true);
        cellsApi.calculateCell(name, sheetName, cellName, options, folder, storage);
    }

    public static void copyCell() throws Exception{
        String name = "CopyCell.xlsx";
        String storage = null;
        String folder = "input";
        String worksheet = "Sheet1";//Source
        String sheetName = "Sheet2";//Destination
        String cellname = "A13";//Source
        String destCellName = "F20";
        int row = 14;//Source row
        int column = 2;//Source column
        cellsApi.copyCell(name, destCellName, sheetName, worksheet, cellname, row, column, folder, storage);
    }

    public static void getCell() throws Exception{
        String name = "GetCell.xlsx";
        String sheetName = "Sheet1";
        String cellOrMethodName = "A2";
        String folder = "input";
        String storage = null;
        File response = cellsApi.getCell(name, sheetName, cellOrMethodName, folder, storage);

    }

    public static void getCells() throws Exception{
        String name = "GetCells.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        int offset = 0;
        int count = 0;
        Cells response = cellsApi.getCells(name, sheetName, offset, count, folder, storage);
    }

    public static void setCellValue() throws Exception{
        String name = "SetCellValue.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        String cellName = "C20";
        String value = "1";
        String type = "String";
        String formula = null;
        Cell response = cellsApi.setCellValue(name, sheetName, cellName, value, type, formula, folder, storage);
    }

    public static void setRangeValue() throws Exception{
        String name = "SetRangeValue.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        String cellarea = "A1:C3";
        String value = "250.5";
        String type = "double";
        cellsApi.setRangeValue(name, sheetName, cellarea, value, type, folder, storage);
    }

    public static void ClearContents() throws Exception{
        String name = "ClearContents.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet4";
        String range = "A1:G19";
        int startRow = 1;
        int startColumn = 1;
        int endRow = 5;
        int endColumn = 8;
        cellsApi.clearContents(name, sheetName, range, startRow, startColumn, endRow, endColumn, folder, storage);
    }

    public static void clearFormats() throws Exception{
        String name = "ClearFormats.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet4";
        String range = "A1:G19";
        int startRow = 1;
        int startColumn = 1;
        int endRow = 5;
        int endColumn = 8;
        cellsApi.clearFormats(name, sheetName, range, startRow, startColumn, endRow, endColumn, folder, storage);
    }

    public static void copyColumns() throws Exception{
        String name = "CopyColumns.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet4";
        int sourceColumnIndex = 2;
        int destinationColumnIndex = 3;
        int columnNumber = 2;
        String worksheet = "Sheet2";
        cellsApi.copyColumns(name, sheetName, sourceColumnIndex, destinationColumnIndex, columnNumber, worksheet, folder, storage);
    }
    public static void deleteColumns() throws Exception{
        String name = "DeleteColumns.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet4";
        Boolean updateReference = true;
        int columnIndex = 1;
        int totalColumns = 2;
        Columns response = cellsApi.deleteColumns(name, sheetName, columnIndex, totalColumns, updateReference,folder, storage);
    }

    public static void getColumn() throws Exception{
        String name = "GetColumn.xlsx";
        String sheetName = "Sheet4";
        int columnIndex = 2;
        String folder = "input";
        String storage = null;
        Column response = cellsApi.getColumn(name, sheetName, columnIndex, folder, storage);
    }

    public static void getColumns() throws Exception{
        String name = "GetColumns.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        Columns response = cellsApi.getColumns(name, sheetName, folder, storage);
    }

    public static void groupColumns() throws Exception{
        String name = "GroupColumns.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet4";
        int firstIndex = 1;
        int lastIndex = 4;
        Boolean hide = true;
        cellsApi.groupColumns(name, sheetName, firstIndex, lastIndex, hide, folder, storage);
    }
    public static void hideColumns() throws Exception{
        String name = "HideColumns.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet4";
        int startColumn = 2;
        int totalColumns = 4;
        cellsApi.hideColumns(name, sheetName, startColumn, totalColumns, folder, storage);
    }

    public static void insertColumns() throws Exception{
        String name = "InsertColumns.xlsx";
        String folder = "inpot";
        String storage = null;
        String sheetName = "Sheet1";
        int columnIndex = 2;
        int columns = 2;
        cellsApi.insertColumns(name, sheetName, columnIndex, columns, folder, storage);
    }

    public static void setColumnStyle() throws Exception{
        String name = "SetColumnStyle.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet4";
        int columnIndex = 2;
        Style style = new Style();
        Font font = new Font();
        font.setIsBold(true);
        font.setIsItalic(true);
        font.setUnderline("Single");
        font.setName("Algerian");
        font.setSize(8);
        style.setFont(font);
        style.setIsTextWrapped(true);
        style.setPattern("Solid");
        style.setPatternColor("red");
        style.setHorizontalAlignment("Center");
        style.setBackgroundColor(new Color(255,0,255,0));
        List<Border> list1= new ArrayList<Border>();
        Border border1 = new Border();
        border1.setLineStyle("Medium");
        border1.setColor(new Color(255, 255, 0, 0));
        border1.setBorderType("EdgeTop");
        list1.add(border1);
        List<Border> list2= new ArrayList<Border>();
        Border border2 = new Border();
        border2.setLineStyle("DashDot");
        border2.setColor(new Color(255, 0, 255, 0));
        border2.setBorderType("EdgeRight");
        list2.add(border2);
        style.setBorderCollection(list1);
        style.setBorderCollection(list2);
        cellsApi.setColumnStyle(name, sheetName, columnIndex, style, folder, storage);
    }

    public static void setColumnWidth() throws Exception{
        String name = "SetColumnWidth.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        int columnIndex = 2;
        Double width =30.5;
        cellsApi.setColumnWidth(name, sheetName, columnIndex, width, folder, storage);
    }

    public static void ungroupColumns() throws Exception{
        String name = "UngroupColumns.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet4";
        int firstIndex = 4;
        int lastIndex = 9;
        cellsApi.ungroupColumns(name, sheetName, firstIndex, lastIndex, folder, storage);
    }

    public static void unhideColumns() throws Exception{
        String name = "UnhideColumns.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet5";
        int startcolumn = 2;
        int totalColumns = 1;
        Double width = null;
        cellsApi.unhideColumns(name, sheetName, startcolumn, totalColumns, width, folder, storage);
    }

    public static void applyRowStyle() throws Exception{
        String name = "ApplyRowStyle.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet2";
        int rowIndex = 3;
        Font font = new Font();
        font.setIsBold(true);
        font.setIsItalic(true);
        font.setUnderline("Single");
        font.setName("Algerian");
        font.setSize(8);
        Style style = new Style();
        style.setFont(font);
        style.setBackgroundColor(new Color(255,0,255,0));
        style.setHorizontalAlignment("Center");
        style.setBorderCollection(new ArrayList<Border>());
        Border border1 = new Border();
        border1.setLineStyle("Medium");
        border1.setColor(new Color(255,255,0,0));
        border1.setBorderType("EdgeTop");
        style.getBorderCollection().add(border1);
        Border border2 = new Border();
        border2.setLineStyle("DashDot");
        border2.setColor(new Color(255,0,255,0));
        border2.setBorderType("EdgeRight");
        style.getBorderCollection().add(border2);
        cellsApi.applyRowStyle(name, sheetName, rowIndex, style, folder, storage);
    }

    public static void copyRows() throws Exception{
        String name = "CopyRows.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet4";
        int sourceRowIndex = 1;
        int destinationRowIndex = 5;
        int rowNumber = 3;
        String worksheet = "Sheet2";
        cellsApi.copyRows(name, sheetName, sourceRowIndex, destinationRowIndex, rowNumber, worksheet, folder, storage);
    }

    public static void deleteRow() throws Exception{
        String name = "DeleteRow.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet2";
        int rowIndex = 1;
        cellsApi.deleteRow(name, sheetName, rowIndex, folder, storage);
    }

    public static void deleteRows() throws Exception{
        String name = "DeleteRows.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet4";
        int rowIndex = 2;
        int totalRows = 2;
        Boolean updateReference = true;//Default value : true
        cellsApi.deleteRows(name, sheetName, rowIndex, totalRows,updateReference, folder, storage);
    }
    public static void getRow() throws Exception{
        String name = "GetRow.xlsx";
        String sheetName = "Sheet4";
        int rowIndex = 2;
        String folder = "input";
        String storage = null;
        Row response = cellsApi.getRow(name, sheetName, rowIndex, folder, storage);
    }

    public static void getRows() throws Exception{
        String name = "GetRows.xlsx";
        String sheetName = "Sheet1";
        String folder = "input";
        String storage = null;
        Rows response = cellsApi.getRows(name, sheetName, folder, storage);
    }

    public static void groupRows() throws Exception{
        String name = "GroupRows.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet4";
        int firstIndex = 2;
        int lastIndex = 6;
        Boolean hide = true;
        cellsApi.groupRows(name, sheetName, firstIndex, lastIndex, hide, folder, storage);
    }

    public static void hideRows() throws Exception{
        String name = "HideRows.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet4";
        int startRow = 4;
        int totalRows = 4;
        cellsApi.hideRows(name, sheetName, startRow, totalRows, folder, storage);
    }

    public static void insertRow() throws Exception{
        String name = "InsertRow.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        int rowIndex = 3;
        cellsApi.insertRow(name, sheetName, rowIndex, folder, storage);
    }

    public static void insertRows() throws Exception{
        String name = "InsertRows.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        int startrow = 5;
        int totalRows = 2;
        String formatAs = "FormatAsBefore";
        cellsApi.insertRows(name, sheetName, startrow, totalRows, formatAs, folder, storage);
    }

    public static void ungroupRows() throws Exception{
        String name = "UngroupRows.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet4";
        int firstIndex = 5;
        int lastIndex = 7;
        Boolean isAll = false;
        cellsApi.ungroupRows(name, sheetName, firstIndex, lastIndex, isAll, folder, storage);
    }

    public static void unhideRows() throws Exception{
        String name = "UnhideRows.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet5";
        int startrow = 15;
        int totalRows = 4;
        Double height = null;
        cellsApi.unhideRows(name, sheetName, startrow, totalRows, height, folder, storage);
    }

    public static void updateRow() throws Exception{
        String name = "UpdateRow.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        int rowIndex = 3;
        Double height = 26.5;
        Row response = cellsApi.updateRow(name, sheetName, rowIndex, height, folder, storage);
    }
}
