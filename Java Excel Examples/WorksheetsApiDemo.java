import spire.cloud.excel.sdk.*;
import spire.cloud.excel.sdk.api.WorksheetsApi;
import spire.cloud.excel.sdk.model.*;

import java.util.ArrayList;

public class WorksheetsApiDemo {
    static String appId = "your id";
    static String appKey = "your key";
    static String baseUrl = "https://api.e-iceblue.cn";
    static Configuration configuration = new Configuration(appId, appKey, baseUrl);
    static WorksheetsApi WorksheetsApi = new WorksheetsApi(configuration);

    public static void autofitRow() throws ApiException {
        String name = "AutofitRow.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        int rowIndex = 1;//Data range: 1~1048576
        int firstColumn = 1;
        int lastColumn = 1;
        WorksheetsApi.autofitRow(name, sheetName, rowIndex, firstColumn, lastColumn, folder, storage);
    }
    public static void autofitRows() throws ApiException {
        String name = "AutofitRows.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        int startRow = 5;
        int endRow = 10;
        boolean onlyAuto = true;//Default value is false. Other values are "false" and "null"
        WorksheetsApi.autofitRows(name, sheetName, startRow, endRow, onlyAuto, folder, storage);
    }
    public static void autofitColumns() throws ApiException {
        String name = "AutofitColumns.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        int firstColumn = 1;
        int lastColumn = 5;
        AutoFitterOptions autoFitterOptions = new AutoFitterOptions(null, true, true, true);
        int firstRow = 3;
        int lastRow = 10;
        WorksheetsApi.autofitColumns(name, sheetName, firstColumn, lastColumn, autoFitterOptions, firstRow, lastRow, folder, storage);
    }
    public static void setBackground() throws ApiException {
        String name = "SetBackground.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        String picPath = "input/SpireXls.png";
        WorksheetsApi.setBackground(name,sheetName,picPath,folder,storage);
    }

    public static void deleteBackground() throws ApiException {
        String name = "DeleteBackground.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        WorksheetsApi.deleteBackground(name, sheetName, folder, storage);
    }
    public static void addComment() throws ApiException {
        String name = "AddComment.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet3";
        String cellName = "C2:D4";
        Comment comment = new Comment();
        comment.setAuthor("jerry");
        comment.setNote("hello");
        comment.setAutoSize(true);
        comment.setTextOrientationType("CounterClockwise");
        comment.setTextVerticalAlignment("Bottom");
        comment.setIsVisible(true);
        WorksheetsApi.addComment(name, sheetName, cellName, comment, folder, storage);
    }
    public static void deleteComment() throws ApiException {
        String name = "DeleteComment.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet2";
        String cellName = "B5";
        WorksheetsApi.deleteComment(name, sheetName, cellName, folder, storage);
    }
    public static void deleteComments() throws ApiException {
        String name = "DeleteComments.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet2";
        WorksheetsApi.deleteComments(name, sheetName, folder, storage);
    }
    public static void getComment() throws ApiException {
        String name = "GetComment.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet3";
        String cellName = "E17";
        WorksheetsApi.getComment(name, sheetName, cellName, folder, storage);
    }
    public static void getComments() throws ApiException {
        String name = "GetComments.xlsx";
        String sheetName = "Sheet2";
        String folder = "input";
        String storage = null;
        WorksheetsApi.getComments(name, sheetName, folder, storage);
    }
    public static void updateComment() throws ApiException {
        String name = "UpdateComment.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet3";
        String cellName = "C6";
        Comment comment = new Comment();
        comment.setAuthor("jerry");
        comment.setNote("Comment");
        comment.setWidth(160);
        comment.setHeight(30);
        comment.setIsVisible(true);
        comment.setTextHorizontalAlignment("center");
        WorksheetsApi.updateComment(name, sheetName, cellName, comment, folder, storage);
    }
    public static void setFreezePanes() throws ApiException {
        String name = "SetFreezePanes.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet2";
        int freezedRows = 3;
        int freezedColumns = 3;
        WorksheetsApi.setFreezePanes(name, sheetName, freezedRows, freezedColumns, folder, storage);
    }
    public static void deleteFreezePanes() throws ApiException {
        String name = "DeleteFreezePanes.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet5";
        WorksheetsApi.deleteFreezePanes(name, sheetName, folder, storage);
    }
    public static void getMergedCell() throws ApiException {
        String name = "GetMergedCell.xlsx";
        String sheetName = "Sheet5";
        int mergedCellIndex = 0;
        String folder = "input";
        String storage = null;
        WorksheetsApi.getMergedCell(name, sheetName, mergedCellIndex, folder, storage);
    }
    public static void getMergedCells() throws ApiException {
        String name = "GetMergedCells.xlsx";
        String sheetName = "Sheet5";
        String folder = "input";
        String storage = null;
        WorksheetsApi.getMergedCells(name, sheetName, folder, storage);
    }
    public static void getNamedRanges() throws ApiException {
        String name = "GetNamedRanges.xlsx";
        String folder = "input";
        String storage = null;
        WorksheetsApi.getNamedRanges(name, folder, storage);
    }
    public static void getNamedRangeValue() throws ApiException {
        String name = "GetNamedRangeValue.xlsx";
        String nameRange = "Sheet1";
        String folder = "input";
        String storage = null;
        WorksheetsApi.getNamedRangeValue(name, nameRange, folder, storage);
    }
    public static void protectWorksheet() throws ApiException {
        String name = "ProtectWorksheet.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet3";
        ProtectSheetParameter protectParameter = new ProtectSheetParameter();
        protectParameter.setPassword("123");
        protectParameter.setProtectionType("scenarios");
        protectParameter.setAllowSelectingLockedCell("true");
        protectParameter.setAllowDeletingColumn("true");
        protectParameter.setAllowFormattingRow("true");
        WorksheetsApi.protectWorksheet(name, sheetName, protectParameter, folder, storage);
    }
    public static void unprotectWorksheet() throws ApiException {
        String name = "UnprotectWorksheet.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        ProtectSheetParameter protectParameter = new ProtectSheetParameter();
        protectParameter.setPassword("123");
        WorksheetsApi.unprotectWorksheet(name, sheetName, protectParameter, folder, storage);
    }
    public static void calculateFormula() throws ApiException {
        String name = "GetCalculateFormula.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet4";
        String formula = "=SUM(D8:D9)";
        WorksheetsApi.calculateFormula(name, sheetName, formula, folder, storage);
    }
    public static void changeVisibilityWorksheet() throws ApiException {
        String name = "ChangeVisibilityWorksheet.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        boolean isVisible = true;
        WorksheetsApi.changeVisibilityWorksheet(name, sheetName, isVisible, folder, storage);
    }
    public static void copyWorksheet() throws ApiException {
        String Destination_name = "CopyWorksheet_destination.xlsx";
        String sourceSheet = "Sheet1";
        String sourceFolder = "input";
        String sourceWorkbook = "CopyWorksheet_original.xlsx";
        String folder = "output";
        String storage = null;
        String sheetName = "Sheet4";
        WorksheetsApi.copyWorksheet(Destination_name, sheetName, sourceSheet, sourceWorkbook, sourceFolder, folder, storage);
    }
    public static void deleteWorksheet() throws ApiException {
        String name = "DeleteWorksheet.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet2";
        WorksheetsApi.deleteWorksheet(name, sheetName, folder, storage);
    }
    public static void getTextItems() throws ApiException {
        String name = "GetTextItems.xlsx";
        String sheetName = "Sheet2";
        String folder = "input";
        String storage = null;
        WorksheetsApi.getTextItems(name, sheetName, folder, storage);
    }
    public static void getWorkSheets() throws ApiException {
        String name = "GetWorkSheets.xlsx";
        String folder = "input";
        String storage = null;
        Worksheets response = WorksheetsApi.getWorkSheets(name, folder, storage);
    }
    public static void getWorkSheetWithFormat() throws ApiException {
        String name = "GetWorkSheetWithFormat.xlsx";
        String sheetName = "Sheet2";
        String format = "html";//Supported formats:html, pdf, png, jpeg, bmp, emf, tiff
        Integer verticalResolution = null;
        Integer horizontalResolution = null;
        String folder = "input";
        String storage = null;
        WorksheetsApi.getWorkSheetWithFormat(name, sheetName, format, verticalResolution, horizontalResolution, folder, storage);
    }
    public static void insertNewWorksheet() throws ApiException {
        String name = "InsertNewWorksheet.xlsx";
        String folder = "input";
        String storage = null;
        String sheetType = "NormalWorksheet";//Available value: NormalWorksheet, ChartSheet
        int index = 2;
        String newSheetName = "SheetInsert";
        WorksheetsApi.insertNewWorksheet(name, newSheetName, index, sheetType, folder, storage);
    }
    public static void moveWorksheet() throws ApiException {
        String name = "MoveWorksheet_After.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet1";
        WorksheetMovingRequest moving = new WorksheetMovingRequest();
        moving.setDestinationWorksheet("sheet4");
        moving.setPosition("after");
        WorksheetsApi.moveWorksheet(name, sheetName, moving, folder, storage);
    }
    public static void renameWorksheet() throws ApiException {
        String name = "RenameWorksheet.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet2";
        String newName = "newSheetName";
        WorksheetsApi.renameWorksheet(name, sheetName, newName, folder, storage);
    }
    public static void replaceText() throws ApiException {
        String name = "ReplaceText.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet4";
        String oldValue = "North America";
        String newValue = "replace";
        boolean matchCase = true;
        WorksheetsApi.replaceText(name, sheetName, oldValue, newValue,matchCase,folder,storage);
    }
    public static void SearchText() throws ApiException {
        String name = "SearchText.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet4";
        String text = "America";
        boolean matchCase = false;
        WorksheetsApi.searchText(name, sheetName, text, matchCase, folder, storage);
    }
    public static void sortRange() throws ApiException {
        String name = "SortRange.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet4";
        String cellArea = "A1:G10";
        DataSorter dataSorter = new DataSorter();
        dataSorter.caseSensitive(true);
        dataSorter.hasHeaders(true);
        dataSorter.sortLeftToRight(false);
        ArrayList<SortKey> KeyList = new ArrayList<SortKey>();
        KeyList.add(new SortKey(0, "Ascending", "Values"));
        KeyList.add(new SortKey(1, "Descending", "Values"));
        dataSorter.setKeyList(KeyList);
        WorksheetsApi.sortRange(name, sheetName, cellArea, dataSorter, folder, storage);
    }
    public static void updateWorksheetProperty() throws ApiException {
        String name = "UpdateWorksheetProperty.xlsx";
        String folder = "input";
        String storage = null;
        String sheetName = "Sheet2";
        Worksheet sheet = new Worksheet();
        sheet.zoom(50);
        sheet.index(3);
        sheet.isGridlinesVisible(false);//Default value is "true".
        sheet.tabColor( new Color(255, 255, 0, 0));
        sheet.isProtected(true);
        sheet.isSelected(true);
        WorksheetsApi.updateWorksheetProperty(name, sheetName, sheet, folder, storage);
    }
    public static void updateWorksheetZoom() throws ApiException {
        String name = "UpdateWorksheetZoom.xlsx";
        String sheetName = "Sheet2";
        int value = 150;
        String folder = "Worksheets/Zoom";
        String storage = null;
        WorksheetsApi.updateWorksheetZoom(name, sheetName, value, folder, storage);
    }
}
