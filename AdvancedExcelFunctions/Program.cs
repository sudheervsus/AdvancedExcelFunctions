using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel=Microsoft.Office.Interop.Excel;

namespace AdvancedExcelFunctions
{
    public class ExcelFunctions
    {
        Excel.Application excelApp = new Excel.Application();
        Excel.Workbook Wb;
        Excel.Worksheet Ws;
        public string OpenExcel(string FilePath, string SheetName = "")
        {
            string result = "";
            try
            {
                excelApp.Visible = true;
                excelApp.DisplayAlerts = false;
                Wb = excelApp.Workbooks.Open(FilePath);
                if (SheetName == "")
                {
                    Ws = (Excel.Worksheet)Wb.ActiveSheet;
                    Ws.Activate();
                }
                else
                {
                    Ws = (Excel.Worksheet)Wb.Worksheets[SheetName];
                    Ws.Activate();
                }
                result = "";
            }
            catch (Exception e)
            {
                result = "Exception caught - " + e.Message;
            }

            return result;
        }

        public string GotoCell(string FilePath, string cellName, string SheetName = "")
        {
            string result = "";
            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Range[cellName].Select();
                    result = "";
                }
            }
            catch (Exception e)
            {
                result = "Exception caught - " + e.Message;
            }

            return result;
        }

        public string GoOneCellRight(string FilePath, string SheetName)
        {
            string result = "";
            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Excel.Range range = (Excel.Range)excelApp.ActiveCell;
                    range.Offset[0, 1].Select();
                    result = "";
                }
            }
            catch (Exception e)
            {
                result = "Exception caught - " + e.Message;
            }

            return result;
        }

        public string GoOneCellLeft(string FilePath, string SheetName)
        {
            string result = "";
            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Excel.Range range = (Excel.Range)excelApp.ActiveCell;
                    if(range.Column !=1)
                    {
                        range.Offset[0, -1].Select();
                    }
                    result = "";
                }
            }
            catch (Exception e)
            {
                result = "Exception caught - " + e.Message;
            }

            return result;
        }

        public string GoOneCellUp(string FilePath, string SheetName)
        {
            string result = "";
            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Excel.Range range = (Excel.Range)excelApp.ActiveCell;
                    if (range.Row != 1)
                    {
                        range.Offset[-1, 0].Select();
                    }
                    result = "";
                }
            }
            catch (Exception e)
            {
                result = "Exception caught - " + e.Message;
            }

            return result;
        }

        public string GoOneCellDown(string FilePath, string SheetName)
        {
            string result = "";
            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Excel.Range range = (Excel.Range)excelApp.ActiveCell;
                    if (range.Row != 1)
                    {
                        range.Offset[1, 0].Select();
                    }
                    result = "";
                }
            }
            catch (Exception e)
            {
                result = "Exception caught - " + e.Message;
            }

            return result;
        }

        public string GoToLastRow(string FilePath, string SheetName)
        {
            string result = "";
            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Excel.Range range = (Excel.Range)excelApp.ActiveCell;
                    range.SpecialCells(XlCellType.xlCellTypeLastCell).Select();
                    result = "";
                }
            }
            catch (Exception e)
            {
                result = "Exception caught - " + e.Message;
            }

            return result;
        }

        public string GoToLastColumn(string FilePath, string SheetName)
        {
            string result = "";
            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Excel.Range range = (Excel.Range)excelApp.ActiveCell;
                    int currentrow = range.Row;
                    Ws.Cells[currentrow, Ws.Columns.Count].End(XlDirection.xlToLeft).Select();
                    result = "";
                }
            }
            catch (Exception e)
            {
                result = "Exception caught - " + e.Message;
            }

            return result;
        }

        public string GetActiveCellAddress(string FilePath, string SheetName)
        {
            string result = "";
            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Excel.Range range = (Excel.Range)excelApp.ActiveCell;
                    result = range.Address[false, false, XlReferenceStyle.xlA1];
                }
            }
            catch (Exception e)
            {
                result = "Exception caught - " + e.Message;
            }

            return result;
        }

        public string GetLastUsedRow(string FilePath, string ColumnName, string SheetName = "")
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    int lastRow = Ws.Range[ColumnName + ":" + ColumnName].Rows.Count;
                    result = lastRow.ToString();
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string GetLastUsedColumn(string FilePath, string SheetName = "")
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    int lastColumn = Ws.UsedRange.Columns.Count;
                    result = lastColumn.ToString();
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string SetCell(string FilePath, string cellName, string cellValue, string SheetName = "")
        {
            string result = "";
            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Range[cellName].Value2 = cellValue;
                    result = "";
                }

            }
            catch (Exception e)
            {
                result = "Exception caught - " + e.Message;
            }

            return result;
        }

        public string SetFormula(string FilePath, string cellName, string Formula, string SheetName = "")
        {
            string result = "";
            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Range[cellName].Formula = Formula;
                    result = "";
                }

            }
            catch (Exception e)
            {
                result = "Exception caught - " + e.Message;
            }

            return result;
        }

        public string RunMacro(string FilePath, string MacroName)
        {
            string result = "";
            try
            {
                if (MacroName == String.Empty)
                {
                    return "Exception caught - Macro Name shouldn't be empty";
                }
                result = OpenExcel(FilePath);
                if (result == "")
                {
                    Ws.GetType().InvokeMember("Run", System.Reflection.BindingFlags.Default | System.Reflection.BindingFlags.InvokeMethod, null, excelApp, new Object[] { MacroName });
                }
            }
            catch (Exception e)
            {
                result = "Exception caught - " + e.Message;
            }
            return result;
        }

        public string AutoFilter(string FilePath, string SheetName, int Field, string FilterRange, string Criteria1, string Criteria2 = "", string Operator = "")
        {
            string result = "";
            try
            {
                object criteria2;
                if (Criteria2 == "")
                {
                    criteria2 = Type.Missing;
                }
                else
                {
                    criteria2 = Criteria2;
                }
                XlAutoFilterOperator filterOperator;
                switch (Operator.ToUpper())
                {
                    case "AND":
                        filterOperator = XlAutoFilterOperator.xlAnd;
                        break;
                    case "OR":
                        filterOperator = XlAutoFilterOperator.xlOr;
                        break;
                    case "TOP10":
                        filterOperator = XlAutoFilterOperator.xlTop10Items;
                        break;
                    case "BOTTOM10":
                        filterOperator = XlAutoFilterOperator.xlBottom10Items;
                        break;
                    default:
                        filterOperator = XlAutoFilterOperator.xlAnd;
                        break;
                }
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Range[FilterRange].AutoFilter(Field, Criteria1, filterOperator, criteria2);
                }

            }
            catch (Exception e)
            {
                result = "Exception caught - " + e.Message;
            }

            return result;
        }

        public string AutoFill(string FilePath, string StartRange, string DestinationRange, string SheetName = "")
        {
            string result = "";
            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Range[StartRange].AutoFill(Ws.Range[DestinationRange]);
                    result = "";
                }
            }
            catch (Exception e)
            {
                result = "Exception caught - " + e.Message;
            }

            return result;
        }

        public string ClearFilter(string FilePath, string SheetName)
        {
            string result = "";
            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    if (Ws.AutoFilterMode)
                    {
                        Ws.AutoFilter.ShowAllData();
                        Ws.AutoFilterMode = false;
                        result = "";
                    }
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string CopyCells(string FilePath, string SheetName, string CopyRange)
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Range[CopyRange].Copy();
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string PasteSpecial(string FilePath, string SheetName, string Range, string PasteType, string PasteSpecialOperation, bool SkipBlanks = false, bool Transpose = false)
        {
            string result = "";

            try
            {
                XlPasteType xlPasteType;
                XlPasteSpecialOperation xlPasteOperation;
                switch (PasteSpecialOperation.ToUpper())
                {
                    case "ADD":
                        xlPasteOperation = XlPasteSpecialOperation.xlPasteSpecialOperationAdd;
                        break;
                    case "SUBTRACT":
                        xlPasteOperation = XlPasteSpecialOperation.xlPasteSpecialOperationSubtract;
                        break;
                    case "MULTIPLY":
                        xlPasteOperation = XlPasteSpecialOperation.xlPasteSpecialOperationMultiply;
                        break;
                    case "DIVIDE":
                        xlPasteOperation = XlPasteSpecialOperation.xlPasteSpecialOperationDivide;
                        break;
                    default:
                        xlPasteOperation = XlPasteSpecialOperation.xlPasteSpecialOperationNone;
                        break;
                }
                switch (PasteType.ToUpper())
                {
                    case "ALL":
                        xlPasteType = XlPasteType.xlPasteAll;
                        break;
                    case "FORMULAS":
                        xlPasteType = XlPasteType.xlPasteFormulas;
                        break;
                    case "VALUES":
                        xlPasteType = XlPasteType.xlPasteValues;
                        break;
                    case "FORMATS":
                        xlPasteType = XlPasteType.xlPasteFormats;
                        break;
                    case "COMMENTS":
                        xlPasteType = XlPasteType.xlPasteComments;
                        break;
                    case "VALIDATION":
                        xlPasteType = XlPasteType.xlPasteValidation;
                        break;
                    default:
                        xlPasteType = XlPasteType.xlPasteAll;
                        break;
                }
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Range[Range].PasteSpecial(xlPasteType, xlPasteOperation, SkipBlanks, Transpose);
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string AddComment(string FilePath, string SheetName, string Range, string CommentText)
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Range[Range].AddComment(CommentText);
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string RefreshAll(string FilePath)
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath, "");
                if (result == "")
                {
                    Wb.RefreshAll();
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string RefreshAllPivotTables(string FilePath)
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath, "");
                if (result == "")
                {
                    foreach(Excel.Worksheet worksheet in Wb.Worksheets)
                    {
                        foreach(Excel.PivotTable pt in worksheet.PivotTables())
                        {
                            pt.RefreshTable();
                            pt.PivotCache().Refresh();
                        }
                    }
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;

        }

        public string RefreshPivotTableByName(string FilePath, string SheetName, string PivotTableName)
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.PivotTables(PivotTableName).RefreshTable();
                    Ws.PivotTables(PivotTableName).PivotCache().Refresh();

                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;

        }

        public string ChangeSourceDataPivotTable(string FilePath, string PivotSheetName, string SourceSheetName, string SourceDataRange, string PivotTableName)
        {
            string result = "";

            try
            {
                string srcData = SourceSheetName + "!" + Ws.Range[SourceDataRange].Address[XlReferenceStyle.xlR1C1];
                result = OpenExcel(FilePath, PivotSheetName);
                if (result == "")
                {
                    Excel.PivotTable pivot = Ws.PivotTables(PivotTableName);
                    pivot.ChangePivotCache(Wb.PivotCaches().Create(XlPivotTableSourceType.xlDatabase, srcData));
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;

        }

        public string InsertSheet(string FilePath, string SheetName)
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath);
                if (result == "")
                {
                    Ws = Wb.Worksheets.Add(After:Wb.Worksheets[Wb.Worksheets.Count],Type:XlSheetType.xlWorksheet);
                    Ws.Name = SheetName;
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string ActivateSheet(string FilePath, string SheetName)
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Activate();
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string RenameSheet(string FilePath, string SheetName, string NewSheetName)
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Name = NewSheetName;
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string HideSheet(string FilePath, string SheetName)
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath);
                if (result == "")
                {
                    Wb.Worksheets[SheetName].Visible = XlSheetVisibility.xlSheetHidden;
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string UnhideSheet(string FilePath, string SheetName)
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath);
                if (result == "")
                {
                    Wb.Worksheets[SheetName].Visible = XlSheetVisibility.xlSheetVisible;
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string DeleteSheet(string FilePath, string SheetName)
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Delete();
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string InsertColumn(string FilePath, string Range, string SheetName = "")
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Range[Range].EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                    result = "";
                }
            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string HideColumn(string FilePath, string SheetName, string ColumnRange)
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Range[ColumnRange].EntireColumn.Hidden = true;
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string UnhideColumn(string FilePath, string SheetName, string ColumnRange)
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Columns[ColumnRange].EntireColumn.Hidden = false;
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string DeleteColumns(string FilePath, string Range, string SheetName = "")
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Range[Range].EntireColumn.Delete(XlDeleteShiftDirection.xlShiftToLeft);
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string InsertRow(string FilePath, string Range, string SheetName = "")
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Range[Range].EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                    result = "";
                }
            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string HideRow(string FilePath, string SheetName, int RowRange)
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Rows[RowRange].EntireRow.Hidden = true;
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string UnhideRow(string FilePath, string SheetName, int RowRange)
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Rows[RowRange].EntireRow.Hidden = false;
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string DeleteRows(string FilePath, string Range, string SheetName = "")
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Range[Range].EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string RemoveDuplicates(string FilePath, string Range, string Columns, string HeaderInfo, string SheetName = "")
        {
            string result = "";

            try
            {
                XlYesNoGuess guess;
                switch (HeaderInfo.ToUpper())
                {
                    case "YES":
                        guess = XlYesNoGuess.xlYes;
                        break;
                    case "NO":
                        guess = XlYesNoGuess.xlNo;
                        break;
                    default:
                        guess = XlYesNoGuess.xlGuess;
                        break;
                }

                string[] cols = Columns.Split(',');
                int[] ArrayofCols = Array.ConvertAll<string, int>(cols, int.Parse);

                result = OpenExcel(FilePath, SheetName);
                if (result == "")
                {
                    Ws.Range[Range].RemoveDuplicates(ArrayofCols, guess);
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string SaveExcel(string FilePath)
        {
            string result = "";

            try
            {
                result = OpenExcel(FilePath, "");
                if (result == "")
                {
                    Wb.Save();
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string SaveAs(string FilePath, string SaveAsFilePath, string SaveAsFileType)
        {
            string result = "";

            try
            {
                XlFileFormat xlFileFormat;
                switch(SaveAsFileType.ToUpper())
                {
                    case "CSV":
                        xlFileFormat = XlFileFormat.xlCSV;
                        break;
                    default:
                        xlFileFormat = XlFileFormat.xlExcel12;
                        break;
                }
                result = OpenExcel(FilePath, "");
                if (result == "")
                {
                    Wb.SaveAs(SaveAsFilePath,xlFileFormat);
                    result = "";
                }

            }
            catch (Exception e)
            {
                return "Exception caught - " + e.Message;
            }

            return result;
        }

        public string CloseExcel(string FilePath, bool SaveChanges)
        {
            string result = "";
            try
            {
                foreach (Excel.Workbook workbook in excelApp.Workbooks)
                {
                    if (workbook.FullName == FilePath)
                    {
                        workbook.Close(SaveChanges);
                        GC.Collect();
                        result = "";
                        break;
                    }
                }
            }catch(Exception e)
            {
                result = "Exception caught - " + e.Message;
            }

            return result;
        }

    }
}
