using DavigoldExcel.Models;
using DavigoldExcel.ViewModel;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace DavigoldExcel.Service
{

    public class Column
    {
        public string Name { get; set; }
        public int Position { get; set; }
    }

    public class ExcelService
    {

        public static List<string> GetColumns()
        {
            List<string> columns = new List<string>();

            // Access the active worksheet
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;

            int lastColumn = GetLastColumnPosition();


            for (int col = 1; col <= lastColumn; col++)
            {
                Range currentRange = (Range)activeSheet.Cells[1, col];
                string cellValue = currentRange.Value2 as string;

                if (currentRange.Comment != null)
                {
                    if(cellValue != null && !String.IsNullOrEmpty(cellValue))
                    {
                        string columnSlug = currentRange.Comment.Text();
                        List<string> columnSlugList = columnSlug.Split(':').ToList();
                        if (columnSlugList.Count == 3)
                        {
                            columns.Add(columnSlugList[2]);
                        }
                        else
                        {
                            columns.Add(columnSlug);
                        }
                    } else
                    {
                        currentRange.ClearComments();
                    }
                }
            }

            return columns;
        }

        public static List<Column> GetColumnsWithPosition()
        {
            List<Column> columns = new List<Column>();

            // Access the active worksheet
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;

            int lastColumn = GetLastColumnPosition();

            for (int col = 1; col <= lastColumn; col++)
            {
                Range currentRange = (Range)activeSheet.Cells[1, col];
                if (currentRange.Comment != null)
                {
                    string columnSlug = currentRange.Comment.Text();
                    List<string> columnSlugList = columnSlug.Split(':').ToList();
                    if (columnSlugList.Count == 3)
                    {
                        columns.Add(new Column() { Name = columnSlugList[2], Position = col });
                    }
                    else
                    {
                        columns.Add(new Column() { Name = columnSlug, Position = col });
                    }
                }
            }

            return columns;
        }

        private static int GetLastColumnPosition()
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;

            // Get the range representing the first row
            Range firstRowRange = (Range)activeSheet.Rows[1];
            Range fullRange = (Range)firstRowRange.Cells[1, activeSheet.Columns.Count];

            // Find the last column with data in the first row
            int lastColumn = fullRange.End[XlDirection.xlToLeft].Column;
            return lastColumn;
        }

        public static void AddColumnAtLastPosition(LabelViewModel Label)
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;

            List<string> CurrentColumns = GetColumns();

            bool isExists = CurrentColumns.Find(column => column == Label.Slug) != null;

            if (isExists) { return; }

            int lastColumn = GetLastColumnPosition();

            string cellValue = (activeSheet.Cells[1, lastColumn] as Range).Value2 as string;

            int insertColumn = !String.IsNullOrWhiteSpace(cellValue) ? lastColumn + 1 : 1;
            Range insertRange = (Range)activeSheet.Cells[1, insertColumn];
            if (insertRange != null)
            {
                if (!String.IsNullOrWhiteSpace(Label.Slug) && !String.IsNullOrEmpty(Label.Name))
                {
                    insertRange.Value = Label.Name;
                    insertRange.ColumnWidth = Label.Name.Length + 3;
                    if (insertRange.Comment != null)
                    {
                        insertRange.ClearComments();
                    }
                    string slug = $"{Label.Module}:{Label.SubModule}:{Label.Slug}";
                    insertRange.AddComment(slug);

                    if (CurrentColumns.Count == 0)
                    {
                        StoreDataInActiveSheet("Module", Label.Module);
                        StoreDataInActiveSheet("SubModule", Label.SubModule);
                    }
                }
            }
        }

        public static void AddKpiAtLastPosition(KpiViewModel Label)
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;

            List<string> CurrentColumns = GetColumns();

            bool isExists = CurrentColumns.Find(column => column == Label.Id.ToString() + "-kpi") != null;

            if (isExists) { return; }

            int lastColumn = GetLastColumnPosition();

            string cellValue = (activeSheet.Cells[1, lastColumn] as Range).Value2 as string;

            int insertColumn = !String.IsNullOrWhiteSpace(cellValue) ? lastColumn + 1 : 1;
            Range insertRange = (Range)activeSheet.Cells[1, insertColumn];
            if (insertRange != null)
            {
                if (!String.IsNullOrWhiteSpace(Label.Id.ToString()) && !String.IsNullOrEmpty(Label.Name))
                {
                    insertRange.Value = Label.Name;
                    insertRange.ColumnWidth = Label.Name.Length + 3;
                    if (insertRange.Comment != null)
                    {
                        insertRange.ClearComments();
                    }
                    insertRange.AddComment(Label.Id.ToString() + "-kpi");

                    if (CurrentColumns.Count == 0)
                    {
                        StoreDataInActiveSheet("Module", Label.Module);
                        StoreDataInActiveSheet("SubModule", Label.SubModule);
                    }
                }
            }
        }

        public static void RemoveColumn(LabelViewModel Label)
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;

            List<string> CurrentColumns = GetColumns();

            bool isExists = CurrentColumns.Find(column => column == Label.Slug) != null;

            if (!isExists) { return; }

            int lastColumn = GetLastColumnPosition();


            for (int col = 1; col <= lastColumn; col++)
            {
                Range CurrentRange = (Range)activeSheet.Cells[1, col];
                if (CurrentRange.Comment != null)
                {
                    string cellValue = CurrentRange.Comment.Text();
                    List<string> cellValueList = cellValue.Split(':').ToList();
                    if (cellValueList.Count == 3)
                    {
                        cellValue = cellValueList[2];
                    }
                    if (cellValue == Label.Slug)
                    {
                        //(activeSheet.Cells[1, col] as Range).Delete();
                        (activeSheet.Columns[col] as Range).Delete();
                        Label.Selection = false;
                        if (CurrentColumns.Count == 1)
                        {
                            StoreDataInActiveSheet("Module", null);
                            StoreDataInActiveSheet("SubModule", null);
                        }
                        break;
                    }
                }
            }
        }

        public static void RemoveKpi(KpiViewModel Label)
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;

            List<string> CurrentColumns = GetColumns();

            bool isExists = CurrentColumns.Find(column => column == Label.Id.ToString() + "-kpi") != null;

            if (!isExists) { return; }

            int lastColumn = GetLastColumnPosition();


            for (int col = 1; col <= lastColumn; col++)
            {
                Range CurrentRange = (Range)activeSheet.Cells[1, col];
                if (CurrentRange.Comment != null)
                {
                    string cellValue = CurrentRange.Comment.Text();
                    if (cellValue == Label.Id.ToString() + "-kpi")
                    {
                        //(activeSheet.Cells[1, col] as Range).Delete();
                        (activeSheet.Columns[col] as Range).Delete();
                        Label.Selection = false;
                        if (CurrentColumns.Count == 1)
                        {
                            StoreDataInActiveSheet("Module", null);
                            StoreDataInActiveSheet("SubModule", null);
                        }
                        break;
                    }
                }
            }
        }

        public static void SyncColumnsComments(ObservableCollection<LabelViewModel> AllLabels, ComboBoxModel SelectedValueType, string downloadValue)
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;


            Range usedRange = activeSheet.UsedRange as Range;

            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    Range currentRange = usedRange.Cells[i, j] as Range;
                    string cellValue = currentRange.Value2 as string;
                    if (currentRange.Comment == null && cellValue != null)
                    {
                        LabelViewModel cellLabel = AllLabels.Where(label => label.Name == cellValue).FirstOrDefault();
                        if (cellLabel != null)
                        {
                            string slug = $"{downloadValue}:{cellLabel.Module}:{cellLabel.SubModule}:{cellLabel.Slug}:{SelectedValueType.Value}";

                            currentRange.AddComment(slug);
                            //currentRange.ColumnWidth = cellValue.Length + 3;
                        }
                    }
                }
            }

            // Get the range representing the first row
            //Range firstRowRange = (Range)activeSheet.Rows[1];
            //Range fullRange = (Range)firstRowRange.Cells[1, activeSheet.Columns.Count];

            //List<string> CurrentColumns = GetColumns();

            //// Find the last column with data in the first row
            //int lastColumn = fullRange.End[XlDirection.xlToLeft].Column;

            //for (int col = 1; col <= lastColumn; col++)
            //{
            //    Range currentRange = (Range)activeSheet.Cells[1, col];
            //    string cellValue = currentRange.Value2 as string;
            //    if (currentRange.Comment == null && cellValue != null)
            //    {
            //        LabelViewModel cellLabel = AllLabels.Where(label => label.Name == cellValue).FirstOrDefault();
            //        if (cellLabel != null)
            //        {
            //            string slug = $"{cellLabel.Module}:{cellLabel.SubModule}:{cellLabel.Slug}";

            //            currentRange.AddComment(slug);
            //            currentRange.ColumnWidth = cellValue.Length + 3;
            //            cellLabel.Selection = true;
            //            if (CurrentColumns.Count == 0)
            //            {
            //                StoreDataInActiveSheet("Module", cellLabel.Module);
            //                StoreDataInActiveSheet("SubModule", cellLabel.SubModule);
            //            }
            //        }
            //    }
            //}
        }

        public static void SyncColumnsKpis(ObservableCollection<KpiViewModel> AllKpis, ComboBoxModel SelectedValueType,string date)
        { 
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;

            Range usedRange = activeSheet.UsedRange;

            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    Range currentRange = usedRange.Cells[i, j] as Range;
                    string cellValue = currentRange.Value2 as string;
                    if (currentRange.Comment == null && cellValue != null)
                    {
                        KpiViewModel cellLabel = AllKpis.Where(label => label.Name == cellValue).FirstOrDefault();
                        if (cellLabel != null)
                        {
                            string comment = $"{cellLabel.Id.ToString()}-kpi:{SelectedValueType.Value}";
                            if (!string.IsNullOrEmpty(date))
                            {
                                comment += ":" + date;
                            }
                            currentRange.AddComment(comment);
                            //currentRange.ColumnWidth = cellValue.Length + 3;
                        }
                    }
                }
            }

            // Get the range representing the first row
            //Range firstRowRange = (Range)activeSheet.Rows[1];
            //Range fullRange = (Range)firstRowRange.Cells[1, activeSheet.Columns.Count];

            //List<string> CurrentColumns = GetColumns();

            //// Find the last column with data in the first row
            //int lastColumn = fullRange.End[XlDirection.xlToLeft].Column;

            //for (int col = 1; col <= lastColumn; col++)
            //{
            //    Range currentRange = (Range)activeSheet.Cells[1, col];
            //    string cellValue = currentRange.Value2 as string;
            //    if (currentRange.Comment == null && cellValue != null)
            //    {
            //        KpiViewModel cellLabel = AllKpis.Where(label => label.Name == cellValue).FirstOrDefault();
            //        if (cellLabel != null)
            //        {
            //            currentRange.AddComment(cellLabel.Id.ToString() + "-kpi");
            //            currentRange.ColumnWidth = cellValue.Length + 3;
            //            cellLabel.Selection = true;
            //            if (CurrentColumns.Count == 0)
            //            {
            //                StoreDataInActiveSheet("Module", cellLabel.Module);
            //                StoreDataInActiveSheet("SubModule", cellLabel.SubModule);
            //            }
            //        }
            //    }
            //}
        }

        public static void StoreDataInActiveSheet(string key, string value)
        {
            Worksheet worksheet = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            // Check if the workbook is not null
            if (worksheet != null)
            {
                // Add or retrieve custom document properties
                if (worksheet.CustomProperties != null)
                {
                    if (worksheet.CustomProperties.Cast<CustomProperty>().Any(p => p.Name == key))
                    {
                        foreach (CustomProperty prop in worksheet.CustomProperties)
                        {
                            if (prop.Name == key)
                            {
                                prop.Delete();
                            }
                        }

                        if (value != null)
                        {
                            worksheet.CustomProperties.Add(key, value);
                        }
                    }
                    else
                    {
                        worksheet.CustomProperties.Add(key, value);
                    }
                }
            }
        }

        public static string GetStoreDataInActiveSheet(string key)
        {
            Worksheet workbook = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;


            // Check if the workbook is not null
            if (workbook != null)
            {
                if (workbook.CustomProperties != null)
                {
                    foreach (CustomProperty prop in workbook.CustomProperties)
                    {
                        if (prop.Name == key)
                        {
                            return prop.Value.ToString();
                        }
                    }
                }
            }

            return null;
        }

        public static string GetStoreDataInWoorkbook(string key)
        {
            Workbook workbook = (Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;

            if (workbook != null)
            {
                // Access custom properties
                DocumentProperties customProperties = (DocumentProperties)workbook.CustomDocumentProperties;

                if (customProperties != null)
                {
                    // Find the custom property by name
                    foreach (DocumentProperty property in customProperties)
                    {
                        if (property.Name == key)
                        {
                            return property.Value.ToString();
                        }
                    }
                }
            }

            return null;
        }

        private static bool IsUrl(string urlString)
        {
            // Try to create a Uri instance from the input string
            if (Uri.TryCreate(urlString, UriKind.Absolute, out Uri uriResult))
            {
                // Check if the Uri scheme is either http or https
                return uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps;
            }

            return false;
        }

        public static void InsertImportedData(List<dynamic> data)
        {
            Application excelApp = Globals.ThisAddIn.Application;
            Worksheet worksheet = (Worksheet)excelApp.ActiveSheet;

            List<Column> CurrentColumns = GetColumnsWithPosition();

            int row = 2; // Start from the second row
            foreach (var item in data)
            {
                foreach (var column in CurrentColumns)
                {
                    Range CurrentRange = worksheet.Cells[row, column.Position] as Range;
                    if (CurrentRange != null)
                    {
                        if (item[column.Name] != null && IsUrl(item[column.Name].ToString()) && column.Name.Contains("download"))
                        {
                            CurrentRange.Value = "Download Document";
                            CurrentRange.Hyperlinks.Add(CurrentRange, item[column.Name]);
                        }
                        else if (item[column.Name] != null && item[column.Name].ToString() != "" && item[column.Name].ToString().Contains("%"))
                        {
                            //    CurrentRange.NumberFormat = "0.00%";
                            CurrentRange.Value = item[column.Name];
                        }
                        else
                        {
                            CurrentRange.Value = item[column.Name];
                        }

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(CurrentRange);
                    }
                }

                row++; // Move to the next row 
            }

        }

        //public static void InsertImportedData(List<dynamic> data, List<dynamic> columns, int startingRow)
        //{
        //    Application excelApp = Globals.ThisAddIn.Application;
        //    Worksheet worksheet = (Worksheet)excelApp.ActiveSheet;

        //    int rowsToInsert = data.Count; // Number of rows you want to insert
        //                                   // Insert rows to push everything down
        //    Range insertRange = worksheet.Rows[startingRow + 1] as Range;
        //    insertRange.Resize[rowsToInsert].Insert(XlInsertShiftDirection.xlShiftDown);

        //    int row = startingRow;
        //    foreach (var item in data)
        //    {
        //        foreach (var column in columns)
        //        {
        //            Range CurrentRange = worksheet.Cells[row, column.Position] as Range;
        //            if (CurrentRange != null)
        //            {
        //                if (item[column.Name] != null && IsUrl(item[column.Name].ToString()) && column.Name.Contains("download"))
        //                {
        //                    CurrentRange.Value = "Download Document";
        //                    CurrentRange.Hyperlinks.Add(CurrentRange, item[column.Name]);
        //                }
        //                else if (item[column.Name] != null && item[column.Name].ToString() != "" && item[column.Name].ToString().Contains("%"))
        //                {
        //                    //    CurrentRange.NumberFormat = "0.00%";
        //                    CurrentRange.Value = item[column.Name];
        //                }
        //                else
        //                {
        //                    CurrentRange.Value = item[column.Name];
        //                }

        //                System.Runtime.InteropServices.Marshal.ReleaseComObject(CurrentRange);
        //            }
        //        }

        //        row++; // Move to the next row 
        //    }

        //}
        public static void InsertImportedData(List<dynamic> data, List<dynamic> columns, int startingRow)
        {
            Application excelApp = Globals.ThisAddIn.Application;
            Worksheet worksheet = (Worksheet)excelApp.ActiveSheet;

            List<string> numericColumns = new List<string>() { "nav", "proceeds", "investment", "ownership", "incremental-invested-eur", "incremental-invested-fx", "incremental-procceds-eur", "incremental-proceeds-fx", "incremental-nav-eur", "incremental-nav-fx", "final-invested-eur", "final-invested-fx", "final-procceds-eur", "final-proceeds-fx", "final-nav-eur", "final-nav-fx" };
            List<string> dateColumns = new List<string>() {"lp-details-date", "date", "formation-on", "started-collaboration-on", "founded-on", "team-joined-on", "team-left-on", "dob", "position-joined-on", "position-left-on", "fund-term-first-closing-date", "fund-term-final-closing-date", "service-management-fee-start-date", "fund-term", "max-fund-term", "investment-period", "max-investment-period", "date-of-liquidation", "closing-start-date", "closing-end-date", "share-date", "received-on", "exit-date", "nda-signed-on", "closing-date" , "capital-table-date", "limit-date", "capitalization-date", "maturity-date", "follow-on-date", "company-asset-status-date", "company-asset-status-closing-date", "investment-date", "due-date", "call-date", "admin-currencies-date", "commitment-date", "recallable-deadline" };

            int rowsToInsert = data.Count; // Number of rows to insert
            int colsCount = columns.Count; // Number of columns

            // Insert rows to make space for new data
            Range insertRange = worksheet.Rows[startingRow + 1] as Range;
            insertRange.Resize[rowsToInsert].Insert(XlInsertShiftDirection.xlShiftDown);

            // Prepare a 2D array to hold data for bulk insert
            object[,] bulkData = new object[rowsToInsert, colsCount];

            // Fill the 2D array with data
            for (int i = 0; i < rowsToInsert; i++)
            {
                var item = data[i];
                for (int j = 0; j < colsCount; j++)
                {
                    var column = columns[j];
                    if (item[column.Name] != null && IsUrl(item[column.Name].ToString()) && column.Name.Contains("download"))
                    {
                        // Insert hyperlink text
                        bulkData[i, j] = "Download Document";

                        // After bulk insert, hyperlinks can be added separately if necessary
                    }
                    else if (item[column.Name] != null && item[column.Name].ToString().Contains("%"))
                    {
                        // Handle percentage formatting if needed
                        bulkData[i, j] = item[column.Name];
                    }
                    else
                    {
                        bulkData[i, j] = item[column.Name];
                    }
                }
            }

            var smallestPosition = columns.Select(col => col.Position).Min();
            var maxPosition = columns.Select(col => col.Position).Max();

            // Define the range where data will be pasted
            Range startCell = worksheet.Cells[startingRow, smallestPosition] as Range;
            Range endCell = worksheet.Cells[(startingRow - 1) + rowsToInsert, maxPosition] as Range;
            Range writeRange = worksheet.Range[startCell, endCell] as Range;

            // Bulk insert data
            writeRange.Value2 = bulkData;

            // Apply number formatting for numeric columns
            foreach (var col in columns)
            {
                int colIndex = col.Position;
                Range formatRange = worksheet.Range[
                    worksheet.Cells[startingRow, colIndex],
                    worksheet.Cells[(startingRow - 1) + rowsToInsert, colIndex]
                ] as Range;

                if (numericColumns.Contains(col.Name))
                {
                    formatRange.NumberFormat = "#,##0.00";
                }
                else if (dateColumns.Contains(col.Name))
                {
                    // Add more keywords if needed for identifying date columns
                    formatRange.NumberFormat = "dd/MM/yyyy";
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(formatRange);
            }

            List<dynamic> hyperlinkColumns = new List<dynamic>();

            for(int i = 0; i < columns.Count; i++)
            {
                if (columns[i].Name.Contains("download") || columns[i].Name.Contains("logo"))
                {
                    hyperlinkColumns.Add(new { Position = columns[i].Position, col = i });
                }
            }

            // Add hyperlinks in a separate loop if any cells require them
            for (int i = 0; i < rowsToInsert; i++)
            {
                for (int j = 0; j < hyperlinkColumns.Count; j++)
                {
                    object currentData = bulkData[i, hyperlinkColumns[j].col];
                    string currentString = currentData.ToString();
                    if (currentString != null && !String.IsNullOrEmpty(currentString) && IsUrl(currentString))
                    {
                        var item = data[i];
                        var Position = hyperlinkColumns[j].Position;
                        // Add hyperlink to the specific cell
                        Range cell = worksheet.Cells[startingRow + i, Position] as Range;
                        string url = item[columns[hyperlinkColumns[j].col].Name];
                        cell.Hyperlinks.Add(cell, url);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(cell);
                    }
                }
            }
        }

    }
}
