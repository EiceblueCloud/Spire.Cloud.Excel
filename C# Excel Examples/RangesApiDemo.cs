using Spire.Cloud.Excel.Sdk.Client;
using Spire.Cloud.Excel.Sdk.Model;
using Spire.Cloud.Excel.Sdk.Api;
using System;
using System.Collections.Generic;
namespace CloudExcel
{
    class RangesApiDemo
    {
        static String appId = "your id";
        static String appKey = "your key";
        static String baseUrl = "https://api.e-iceblue.cn";
        static Configuration configuration = new Configuration(appId, appKey, baseUrl);
        static RangesApi rangesApi = new RangesApi(configuration);
        public static void CopyRanges()
        {
            string name = "CopyRanges.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            RangeCopyRequest rangeOperate = new RangeCopyRequest();
            rangeOperate.Operate = "copystyle";
            Range range = new Range();
            range.FirstColumn = 1;
            range.FirstRow = 1;
            range.RowCount = 10;
            range.ColumnCount = 10;
            Range range2 = new Range();
            range2.FirstRow = 30;
            range2.FirstColumn = 1;
            range2.RowCount = 5;
            range2.ColumnCount = 8;
            rangeOperate.Source = range;
            rangeOperate.Target = range2;
            rangesApi.CopyRanges(name, sheetName, rangeOperate, folder, storage);
        }

        public static void GetRangeValue()
        {
            string name = "GetRangeValue.xlsx";
            string sheetName = "Sheet1";
            string folder = "input";
            string storage = null;
            string namerange = "C1:D3";
            var response = rangesApi.GetRangeValue(name, sheetName, namerange, null, null, null, null, folder, storage);
        }

        public static void MergeRange()
        {
            string name = "MergeRange.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            Range range = new Range();
            range.FirstRow = 3;
            range.FirstColumn = 3;
            range.RowCount = 5;
            range.ColumnCount = 8;
            rangesApi.MergeRange(name, sheetName, range, folder, storage);
        }

        public static void MoveRange()
        {
            string name = "MoveRange.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            Range range = new Range();
            range.FirstRow = 1;
            range.FirstColumn = 1;
            range.ColumnCount = 1;
            range.RowCount = 1;
            int? destRow = 25;
            int? destColumn = 3; 
            rangesApi.MoveRange(name, sheetName, range, destRow, destColumn, folder, storage);
        }
        public static void SetRangeColumnWidth()
        {
            string name = "SetRangeColumnWidth.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            Range range = new Range();
            range.FirstRow = 1;
            range.FirstColumn = 1;
            range.ColumnCount = 5;
            range.RowCount = 2;
            double? value = 15;
            rangesApi.SetRangeColumnWidth(name, sheetName, range, value, folder, storage);
        }
        public static void SetRangeOutlineBorder()
        {
            string name = "SetRangeOutlineBorder.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            RangeSetOutlineBorderRequest rangeOperate =
                new RangeSetOutlineBorderRequest();
            rangeOperate.BorderEdge = "outline";
            rangeOperate.BorderStyle = "Medium";
            rangeOperate.BorderColor = new Color(255, 255, 0, 0);
            Range range = new Range();
            range.FirstColumn = 1;
            range.FirstRow = 1;
            range.ColumnCount = 10;
            range.RowCount = 5;
            rangeOperate.Range = range;
            rangesApi.SetRangeOutlineBorder(name, sheetName, rangeOperate, folder, storage);
        }

        public static void SetRangeRowHeight()
        {
            string name = "SetRangeRowHeight.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            Range range = new Range();
            range.FirstRow = 1;
            range.FirstColumn = 1;
            range.ColumnCount = 5;
            range.RowCount = 10;
            double? value = 20;
            rangesApi.SetRangeRowHeight(name, sheetName, range, value, folder, storage);
        }

        public static void SetRangeStyle()
        {
            string name = "SetRangeStyle_1.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            RangeSetStyleRequest rangeOperate = new RangeSetStyleRequest();
            rangeOperate.Style = new Style();
            rangeOperate.Style.Font = new Font();
            rangeOperate.Style.Font.Color = new Color(255, 255, 0, 0);
            rangeOperate.Style.Font.Underline = "single";
            rangeOperate.Style.Font.Size = 12;
            rangeOperate.Style.Font.Name = "楷体";
            rangeOperate.Style.Font.IsItalic = true;
            rangeOperate.Style.Font.IsStrikeout = true;
            rangeOperate.Style.Font.IsBold = true;
            rangeOperate.Style.BorderCollection = new List<Border>();
            rangeOperate.Style.BorderCollection.Add(new Border("Medium", new Color(255, 255, 0, 0), "EdgeTop"));
            rangeOperate.Style.BorderCollection.Add(new Border("DashDot", new Color(255, 0, 255, 0), "EdgeRight"));
            rangeOperate.Style.BorderCollection.Add(new Border("Thin", new Color(0, 160, 32, 240), "EdgeLeft"));
            rangeOperate.Style.HorizontalAlignment = "left";
            rangeOperate.Style.IndentLevel = 1;
            rangeOperate.Style.TextDirection = "RightToLeft";
            rangeOperate.Style.BackgroundColor = new Color(0, 255, 255, 0);
            rangeOperate.Style.IsTextWrapped = true;
            Range range = new Range();
            range.FirstColumn = 2;
            range.FirstRow = 1;
            range.ColumnCount = 1;
            range.RowCount = 10;
            rangeOperate.Range = range;
            rangesApi.SetRangeStyle(name, sheetName, rangeOperate, folder, storage);
        }

        public static void SetRangeValue()
        {
            string name = "SetRangeValue.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet2";
            Range range = new Range();
            range.FirstRow = 1;
            range.FirstColumn = 1;
            range.ColumnCount = 5;
            range.RowCount = 2;
            string value = "hello,冰蓝";
            bool? isConverted = true;
            bool? setStyle = false;
            rangesApi.SetRangeValue(name, sheetName, range, value, isConverted, setStyle, folder, storage);
        }

        public static void UnMergeRange()
        {
            string name = "UnMergeRange.xlsx";
            string folder = "input";
            string storage = null;
            string sheetName = "Sheet1";
            Range range = new Range();
            range.FirstRow = 3;
            range.FirstColumn = 4;
            range.RowCount = 2;
            range.ColumnCount = 5;
            rangesApi.UnMergeRange(name, sheetName, range, folder, storage);
        }
    }
}
