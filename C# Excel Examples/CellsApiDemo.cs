using Spire.Cloud.Excel.Sdk.Api;
using Spire.Cloud.Excel.Sdk.Client;
using Spire.Cloud.Excel.Sdk.Model;
using System;
using System.Collections.Generic;

namespace CloudExcel
{
    class CellsApiDemo
    {
        static String appId = "your id";
        static String appKey = "your key";
        static String baseUrl = "https://api.e-iceblue.cn";
        static Configuration configuration = new Configuration(appId, appKey, baseUrl);
        static CellsApi cellsApi = new CellsApi(configuration);
        public static void MergeCells()
        {       
            string name = "MergeCells.xlsx";
            string folder = "input";
            string storage = null;

            string sheetName = "Sheet4";
            int? startRow = 2;
            int? startColumn = 3;
            int? totalRows = 4;
            int? totalColumns = 4;
            cellsApi.MergeCells(name, sheetName, startRow, startColumn, totalRows, totalColumns, folder, storage);
        }

        public static void UnMergeCells()
        {
            string name = "UnMergeCells.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet2";
            int? startRow = 3;
            int? startColumn = 1;
            int? totalRows = 1;
            int? totalColumns = 2;
            cellsApi.UnMergeCells(name, sheetName, startRow, startColumn, totalRows, totalColumns, folder, storage);
        }

        public static void GetCellStyle()
        {          
            string name = "GetCellStyle.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet4";
            string cellName = "A4";
            var response = cellsApi.GetCellStyle(name, sheetName, cellName, folder, storage);
        }

        public static void SetCellStyle()
        {
            string name = "SetCellStyle.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet4";
            string cellName = "C8";
            Style style = new Style();
            Font font = new Font();
            font.IsBold = true;
            font.IsItalic = true;
            font.Underline = "Single";
            font.Name = "Algerian";
            font.Size = 8;
            style.Font = font;
            style.BackgroundColor = new Color(255, 0, 255, 0);
            style.HorizontalAlignment = "Center";
            style.BorderCollection = new List<Border>();
            style.BorderCollection.Add(new Border("Medium", new Color(255, 255, 0, 0), "EdgeTop"));
            style.BorderCollection.Add(new Border("DashDot", new Color(255, 0, 255, 0), "EdgeRight"));
            var response = cellsApi.SetCellStyle(name, sheetName, cellName, style, folder, storage);
        }

        public static void CalculateCell()
        {       
            string name = "CalculateCell.xlsx";
            string storage = null;
            string folder = "input";
            string sheetName = "Sheet5";
            string cellName = "G6";
            CalculationOptions options = new CalculationOptions();
            options.CalcStackSize = 0;
            options.PrecisionStrategy = "";
            options.IgnoreError = true;
            options.Recursive = true;
            cellsApi.CalculateCell(name, sheetName, cellName,options, folder, storage);
        }

        public static void CopyCell()
        {
            string name = "CopyCell.xlsx";
            string storage = null;
            string folder = "input";
            string worksheet = "Sheet1";//Source
            string sheetName = "Sheet2";//Destination
            string cellname = "A13";//Source
            string destCellName = "F20";
            int? row = 14;//Source row
            int? column = 2;//Source column
            cellsApi.CopyCell(name, destCellName, sheetName, worksheet, cellname, row, column, folder, storage);
        }

        public static void GetCell()
        {
            string name = "GetCell.xlsx";
            string sheetName = "Sheet1";
            string cellOrMethodName = "A2";
            string folder = "input";
            string storage = null;
            var response = cellsApi.GetCell(name, sheetName, cellOrMethodName, folder, storage);

        }

        public static void GetCells()
        {
            string name = "GetCells.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            int? offset = 0;
            int? count = 0;
            var response = cellsApi.GetCells(name, sheetName, offset, count, folder, storage);
        }

        public static void SetCellValue()
        {
            string name = "SetCellValue.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            string cellName = "C20";
            string value = "1";
            string type = "string";
            string formula = null;
            var response = cellsApi.SetCellValue(name, sheetName, cellName, value, type, formula, folder, storage);
        }

        public static void SetRangeValue()
        {
            string name = "SetRangeValue.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            string cellarea = "A1:C3";
            string value = "250.5";
            string type = "double";
            cellsApi.SetRangeValue(name, sheetName, cellarea, value, type, folder, storage);
        }

        public static void ClearContents()
        {
            string name = "ClearContents.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet4";
            string range = "A1:G19";
            int? startRow = null;
            int? startColumn = null;
            int? endRow = null;
            int? endColumn = null;
            cellsApi.ClearContents(name, sheetName, range, startRow, startColumn, endRow, endColumn, folder, storage);
        }

        public static void ClearFormats()
        {
            string name = "ClearFormats.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet4";
            string range = "A1:G19";
            int? startRow = null;
            int? startColumn = null;
            int? endRow = null;
            int? endColumn = null;
            cellsApi.ClearFormats(name, sheetName, range, startRow, startColumn, endRow, endColumn, folder, storage);
        }

        public static void CopyColumns()
        {
            string name = "CopyColumns.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet4";
            int? sourceColumnIndex = 2;
            int? destinationColumnIndex = 3;
            int? columnNumber = 2;
            string worksheet = "Sheet2";
            cellsApi.CopyColumns(name, sheetName, sourceColumnIndex, destinationColumnIndex, columnNumber, worksheet, folder, storage);
        }
        public static void DeleteColumns()
        { 
            string name = "DeleteColumns.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet4";
            int? columnIndex = 1;
            int? totalColumns = 2;
            var response = cellsApi.DeleteColumns(name, sheetName, columnIndex, totalColumns, folder, storage);
        }

        public static void GetColumn()
        {
            string name = "GetColumn.xlsx";
            string sheetName = "Sheet4";
            int? columnIndex = 2;
            string folder = "input";
            string storage = null;
            var response = cellsApi.GetColumn(name, sheetName, columnIndex, folder, storage);
        }

        public static void GetColumns()
        {
            string name = "GetColumns.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            var response = cellsApi.GetColumns(name, sheetName, folder, storage);
        }

        public static void GroupColumns()
        {
            string name = "GroupColumns.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet4";
            int? firstIndex = 1;
            int? lastIndex = 4;
            bool? hide = true;
            cellsApi.GroupColumns(name, sheetName, firstIndex, lastIndex, hide, folder, storage);
        }
        public static void HideColumns()
        {
            string name = "HideColumns.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet4";
            int? startColumn = 2;
            int? totalColumns = 4;
            cellsApi.HideColumns(name, sheetName, startColumn, totalColumns, folder, storage);
        }

        public static void InsertColumns()
        {
            string name = "InsertColumns.xlsx";
            string folder = "inpot";
            string storage = null;
            string sheetName = "Sheet1";
            int? columnIndex = 2;
            int? columns = 2;
            string formatAs = "FormatAsBefore";
            cellsApi.InsertColumns(name, sheetName, columnIndex, columns, formatAs, folder, storage);
        }

        public static void SetColumnStyle()
        {
            string name = "SetColumnStyle.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet4";
            int? columnIndex = 2;
            Style style = new Style();
            Font font = new Font();
            font.IsBold = true;
            font.IsItalic = true;
            font.Underline = "Single";
            font.Name = "Algerian";
            font.Size = 8;
            style.Font = font;
            style.IsTextWrapped = true;
            style.Pattern = "Solid";
            style.PatternColor = "red";
            style.HorizontalAlignment = "Center";
            style.BackgroundColor = new Color(255, 0, 255, 0);
            style.BorderCollection = new List<Border>();
            style.BorderCollection.Add(new Border("Medium", new Color(255, 255, 0, 0), "EdgeTop"));
            style.BorderCollection.Add(new Border("DashDot", new Color(255, 0, 255, 0), "EdgeRight"));
            cellsApi.SetColumnStyle(name, sheetName, columnIndex, style, folder, storage);
        }

        public static void SetColumnWidth()
        {
            string name = "SetColumnWidth.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            int? columnIndex = 2;
            double? width = 40;
            cellsApi.SetColumnWidth(name, sheetName, columnIndex, width, folder, storage);
        }

        public static void UngroupColumns()
        {
            string name = "UngroupColumns.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet4";
            int? firstIndex = 4;
            int? lastIndex = 9;
            cellsApi.UnGroupColumns(name, sheetName, firstIndex, lastIndex, folder, storage);
        }

        public static void UnhideColumns()
        {
            string name = "UnhideColumns.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet5";
            int? startcolumn = 2;
            int? totalColumns = 1;
            double? width = null;
            cellsApi.UnHideColumns(name, sheetName, startcolumn, totalColumns, width, folder, storage);
        }

        public static void ApplyRowStyle()
        {
            string name = "ApplyRowStyle.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet2";
            int? rowIndex = 3;
            Style style = new Style();
            Font font = new Font();
            font.IsBold = true;
            font.IsItalic = true;
            font.Underline = "Single";
            font.Name = "Algerian";
            font.Size = 8;
            style.Font = font;
            style.BackgroundColor = new Color(255, 0, 255, 0);
            style.HorizontalAlignment = "Center";
            style.BorderCollection = new List<Border>();
            style.BorderCollection.Add(new Border("Medium", new Color(255, 255, 0, 0), "EdgeTop"));
            style.BorderCollection.Add(new Border("DashDot", new Color(255, 0, 255, 0), "EdgeRight"));
            cellsApi.ApplyRowStyle(name, sheetName, rowIndex, style, folder, storage);
        }

        public static void CopyRows()
        {
            string name = "CopyRows.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet4";
            int? sourceRowIndex = 1;
            int? destinationRowIndex = 5;
            int? rowNumber = 3;
            string worksheet = "Sheet2"; 
            cellsApi.CopyRows(name, sheetName, sourceRowIndex, destinationRowIndex, rowNumber, worksheet, folder, storage);
        }

        public static void DeleteRow()
        {
            string name = "DeleteRow.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet2";
            int? rowIndex = 1;
            cellsApi.DeleteRow(name, sheetName, rowIndex, folder, storage);
        }

        public static void DeleteRows()
        {
            string name = "DeleteRows.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet4";
            int? rowIndex = 2;
            int? totalRows = 2;
            cellsApi.DeleteRows(name, sheetName, rowIndex, totalRows, folder, storage);
        }
        public static void GetRow()
        {
            string name = "GetRow.xlsx";
            string sheetName = "Sheet4";
            int? rowIndex = 2;
            string folder = "input";
            string storage = null;
            var response = cellsApi.GetRow(name, sheetName, rowIndex, folder, storage);
        }

        public static void GetRows()
        {
            string name = "GetRows.xlsx";
            string sheetName = "Sheet1";
            string folder = "input";
            string storage = null;
            var response = cellsApi.GetRows(name, sheetName, folder, storage);
        }

        public static void GroupRows()
        {
            string name = "GroupRows.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet4";
            int? firstIndex = 2;
            int? lastIndex = 6;
            bool? hide = true;
            cellsApi.GroupRows(name, sheetName, firstIndex, lastIndex, hide, folder, storage);
        }

        public static void HideRows()
        {
            string name = "HideRows.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet4";
            int? startrow = 4;
            int? totalRows = 4;
            cellsApi.HideRows(name, sheetName, startrow, totalRows, folder, storage);
        }

        public static void InsertRow()
        {
            string name = "InsertRow.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            int? rowIndex = 3;
            cellsApi.InsertRow(name, sheetName, rowIndex, folder, storage);
        }

        public static void InsertRows()
        {
            string name = "InsertRows.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            int? startrow = 5;
            int? totalRows = 2;
            string formatAs = "FormatAsBefore";
            cellsApi.InsertRows(name, sheetName, startrow, totalRows, formatAs, folder, storage);
        }

        public static void UngroupRows()
        {         
            string name = "UngroupRows.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet4";
            int? firstIndex = 5;
            int? lastIndex = 7;
            bool? isAll = false;
            cellsApi.UnGroupRows(name, sheetName, firstIndex, lastIndex, isAll, folder, storage);
        }

        public static void UnhideRows()
        {
            string name = "UnhideRows.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet5";
            int? startrow = 15;
            int? totalRows = 4;
            double? height = null;
            cellsApi.UnHideRows(name, sheetName, startrow, totalRows, height, folder, storage);
        }

        public static void UpdateRow()
        { 
            string name = "UpdateRow.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            int? rowIndex = 3;
            double? height = 26;
            var response = cellsApi.UpdateRow(name, sheetName, rowIndex, height, folder, storage);
        }
    }
}
