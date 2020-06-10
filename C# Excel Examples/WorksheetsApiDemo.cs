using Spire.Cloud.Excel.Sdk.Api;
using Spire.Cloud.Excel.Sdk.Client;
using Spire.Cloud.Excel.Sdk.Model;
using System;
using System.Collections.Generic;
using System.IO;

namespace CloudExcel
{
    class WorksheetsApiDemo
    {
        static String appId = "your id";
        static String appKey = "your key";
        static String baseUrl = "https://api.e-iceblue.cn";
        static Configuration configuration = new Configuration(appId, appKey, baseUrl);
        static WorksheetsApi WorksheetsApi = new WorksheetsApi(configuration);
        public static void AutofitRow()
        {
            string name = "AutofitRow.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            int? rowIndex = 1;//Data range: 1~1048576
            int? firstColumn = 1;
            int? lastColumn = 1;
            WorksheetsApi.AutofitRow(name, sheetName, rowIndex, firstColumn, lastColumn, folder, storage);
        }
        public static void AutofitRows()
        {
            string name = "AutofitRows.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            int? startRow = 5;
            int? endRow = 10;
            bool? onlyAuto = true;//other values are "false" and "null"
            WorksheetsApi.AutofitRows(name, sheetName, startRow, endRow, onlyAuto, folder, storage);
        }
        public static void AutofitColumns()
        {           
            string name = "AutofitColumns.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            int? firstColumn = 1;
            int? lastColumn = 5;
            AutoFitterOptions autoFitterOptions = new AutoFitterOptions(true, true, true);
            int? firstRow = 3;
            int? lastRow = 10;
            WorksheetsApi.AutofitColumns(name, sheetName, firstColumn, lastColumn, autoFitterOptions, firstRow, lastRow, folder, storage);
        }
        public static void SetBackground()
        {  
            string name = "SetBackground.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            using (Stream picFile = new FileStream(Configuration.RootPath + @"..\..\inputFile\SpireXls.png", FileMode.Open))
            WorksheetsApi.SetBackground(name, sheetName, null, folder, picFile, storage);
        }

        public static void DeleteBackground()
        {          
            string name = "DeleteBackground.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            WorksheetsApi.DeleteBackground(name, sheetName, folder, storage);
        }
        public static void AddComment()
        { 
            string name = "AddComment.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet3";
            string cellName = "C2:D4";
            Comment comment = new Comment();
            comment.Author = "jerry";
            comment.Note = "hello";
            comment.AutoSize = true;
            comment.TextOrientationType = " CounterClockwise";
            comment.TextVerticalAlignment = "Bottom";
            comment.IsVisible = true;
            WorksheetsApi.AddComment(name, sheetName, cellName, comment, folder, storage);
        }
        public static void DeleteComment()
        {
            string name = "DeleteComment.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet2";
            string cellName = "B5";
            WorksheetsApi.DeleteComment(name, sheetName, cellName, folder, storage);
        }
        public static void DeleteComments()
        {
            string name = "DeleteComments.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet2";
            WorksheetsApi.DeleteComments(name, sheetName, folder, storage);
        }
        public static void GetComment()
        {
            string name = "GetComment.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet3";
            string cellName = "E17";
            WorksheetsApi.GetComment(name, sheetName, cellName, folder, storage);
        }
        public static void GetComments()
        {            
            string name = "GetComments.xlsx";
            string sheetName = "Sheet2";
            string folder = "input";
            string storage = null;
            WorksheetsApi.GetComments(name, sheetName, folder, storage);
        }
        public static void UpdateComment()
        {
            string name = "UpdateComment.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet3";
            string cellName = "C6";
            Comment comment = new Comment();
            comment.Author = "jerry";
            comment.Note = "Hello,Comments";
            comment.Width = 160;
            comment.Height = 30;
            comment.IsVisible = true;
            comment.TextHorizontalAlignment = "center";
            WorksheetsApi.UpdateComment(name, sheetName, cellName, comment, folder, storage);
        }
        public static void SetFreezePanes()
        {
            string name = "SetFreezePanes.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet2";
            int? freezedRows = 3;
            int? freezedColumns = 3;
            WorksheetsApi.SetFreezePanes(name, sheetName, freezedRows, freezedColumns, folder, storage);
        }
        public static void DeleteFreezePanes()
        {
            string name = "DeleteFreezePanes.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet5";
            WorksheetsApi.DeleteFreezePanes(name, sheetName, folder, storage);
        }
        public static void GetMergedCell()
        {
            string name = "GetMergedCell.xlsx";
            string sheetName = "Sheet5";
            int? mergedCellIndex = 0;
            string folder = "input";
            string storage = null;
            WorksheetsApi.GetMergedCell(name, sheetName, mergedCellIndex, folder, storage);
        }
        public static void GetMergedCells()
        {
            string name = "GetMergedCells.xlsx";
            string sheetName = "Sheet5";
            string folder = "input";
            string storage = null;
            WorksheetsApi.GetMergedCells(name, sheetName, folder, storage);
        }
        public static void GetNamedRanges()
        {          
            string name = "GetNamedRanges.xlsx";
            string folder = "input";
            string storage = null;
            WorksheetsApi.GetNamedRanges(name, folder, storage);
        }
        public static void GetNamedRangeValue()
        {
            string name = "GetNamedRangeValue.xlsx";
            string namerange = "Sheet1";
            string folder = "input";
            string storage = null;
            WorksheetsApi.GetNamedRangeValue(name, namerange, folder, storage);
        }
        public static void ProtectWorksheet()
        {
            string name = "ProtectWorksheet.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet3";
            ProtectSheetParameter protectParameter = new ProtectSheetParameter();
            protectParameter.Password = "123";
            protectParameter.ProtectionType = "scenarios";
            protectParameter.AllowSelectingLockedCell = "true";
            protectParameter.AllowDeletingColumn = "true";
            protectParameter.AllowFormattingRow = "true";
            WorksheetsApi.ProtectWorksheet(name, sheetName, protectParameter, folder, storage);
        }
        public static void UnprotectWorksheet()
        {
            string name = "UnprotectWorksheet.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            ProtectSheetParameter protectParameter = new ProtectSheetParameter();
            protectParameter.Password = "123";
            WorksheetsApi.UnProtectWorksheet(name, sheetName, protectParameter, folder, storage);
        }
        public static void CalculateFormula()
        {
            string name = "GetCalculateFormula.xlsx";
            string folder = "ExcelDocument";
            string storage = null;
            string sheetName = "Sheet4";
            string formula = "=SUM(D8:D9)";
            WorksheetsApi.CalculateFormula(name, sheetName, formula, folder, storage);
        }
        public static void ChangeVisibilityWorksheet()
        {
            string name = "ChangeVisibilityWorksheet.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            bool? isVisible = true;
            WorksheetsApi.ChangeVisibilityWorksheet(name, sheetName, isVisible, folder, storage);
        }
        public static void CopyWorksheet()
        {
            string Destination_name = "CopyWorksheet_destination.xlsx";
            string sourceSheet = "Sheet1";
            string sourceFolder = "input";
            string sourceWorkbook = "CopyWorksheet_original.xlsx";
            string folder = "output";
            string storage = null;
            string sheetName = "Sheet4";
            WorksheetsApi.CopyWorksheet(Destination_name, sheetName, sourceSheet, sourceWorkbook, sourceFolder, folder, storage);
        }
        public static void DeleteWorksheet()
        {
            string name = "DeleteWorksheet.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet2";
            WorksheetsApi.DeleteWorksheet(name, sheetName, folder, storage);
        }
        public static void GetTextItems()
        {
            string name = "GetTextItems.xlsx";
            string sheetName = "Sheet2";
            string folder = "input";
            string storage = null;
            WorksheetsApi.GetTextItems(name, sheetName, folder, storage);
        }
        public static void GetWorkSheets()
        {
            string name = "GetWorkSheets.xlsx";
            string folder = "input";
            string storage = null;
            WorksheetsApi.GetWorkSheets(name, folder, storage);
        }
        public static void GetWorkSheetWithFormat()
        {
            string name = "GetWorkSheetWithFormat.xlsx";
            string sheetName = "Sheet2";
            string format = "html";//supports formats:html/pdf/png/jpeg/bmp/emf/tiff
            int? verticalResolution = null;
            int? horizontalResolution = null;
            string folder = "input";
            string storage = null;
            WorksheetsApi.GetWorkSheetWithFormat(name, sheetName, format, verticalResolution, horizontalResolution, folder, storage);
        }
        public static void InsertNewWorksheet()
        {
            string name = "InsertNewWorksheet.xlsx";
            string folder = "input";
            string storage = null;
            string sheettype = "NormalWorksheet";
            int? index = 2;
            string newsheetname = "SheetInsert";
            WorksheetsApi.InsertNewWorksheet(name, newsheetname, index, sheettype, folder, storage);
        }
        public static void MoveWorksheet()
        {
            string name = "MoveWorksheet_After.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            WorksheetMovingRequest moving = new WorksheetMovingRequest();
            moving.DestinationWorksheet = "sheet4";
            moving.Position = "after";
            WorksheetsApi.MoveWorksheet(name, sheetName, moving, folder, storage);
        }
        public static void RenameWorksheet()
        {
            string name = "RenameWorksheet.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet2";
            string newname = "newSheetName";
            WorksheetsApi.RenameWorksheet(name, sheetName, newname, folder, storage);
        }
        public static void ReplaceText()
        {
            string name = "ReplaceText.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet4";
            string oldValue = "North America";
            string newValue = "replace";
            bool matchCase = true;
            WorksheetsApi.ReplaceText(name, sheetName, oldValue, newValue, matchCase, folder, storage);
        }
        public static void SearchText()
        {
            string name = "SearchText.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet4";
            string text = "America";
            bool matchCase = false;
            WorksheetsApi.SearchText(name, sheetName, text, matchCase, folder, storage);
        }
        public static void SortRange()
        {
            string name = "SortRange.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet4";
            string cellArea = "A1:G10";
            DataSorter dataSorter = new DataSorter();
            dataSorter.CaseSensitive = true;
            dataSorter.HasHeaders = true;
            dataSorter.SortLeftToRight = false;
            dataSorter.KeyList = new List<SortKey>();
            dataSorter.KeyList.Add(new SortKey(0, "Ascending", "Values"));
            dataSorter.KeyList.Add(new SortKey(1, "Descending", "Values"));
            WorksheetsApi.SortRange(name, sheetName, cellArea, dataSorter, folder, storage);
        }
        public static void UpdateWorksheetProperty()
        {
            string name = "UpdateWorksheetProperty.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet2";
            Worksheet sheet = new Worksheet();
            sheet.Zoom = 50;
            sheet.Index = 3;
            sheet.IsGridlinesVisible = false;//default value is "true".
            sheet.TabColor = new Color(255, 255, 0, 0);
            sheet.IsProtected = true;
            sheet.IsSelected = true;
            WorksheetsApi.UpdateWorksheetProperty(name, sheetName, sheet, folder, storage);
        }
        public static void UpdateWorksheetZoom()
        {
            string name = "UpdateWorksheetZoom.xlsx";
            string sheetName = "Sheet2";
            int? value = 150;
            string folder = "Worksheets/Zoom";
            string storage = null;
            WorksheetsApi.UpdateWorksheetZoom(name, sheetName, value, folder, storage);
        }
    }
}
