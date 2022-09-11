//Created by Alexander Fields https://github.com/roku674

using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelCSharp
{
    /// <summary>
    /// Call this class to be able to open workbooks
    /// </summary>
    public class Excel
    {
        private static Dictionary<string, Microsoft.Office.Interop.Excel.Worksheet> dict = new Dictionary<string, Microsoft.Office.Interop.Excel.Worksheet>();
        private readonly string path = "";
        private Microsoft.Office.Interop.Excel._Application excel = new Application();
        private Microsoft.Office.Interop.Excel.Workbook wb;
        private Microsoft.Office.Interop.Excel.Worksheet ws;

        /// <summary>
        /// Excel Constructor
        /// </summary>
        /// <param name="path">Include the file extension</param>
        /// <param name="Sheet">Sheet's start at page 1 not page 0</param>
        public Excel(string path, int Sheet)
        {
            this.path = path;

            if (!System.IO.File.Exists(path))
            {
                Application app = new Microsoft.Office.Interop.Excel.Application();
                app.DisplayAlerts = false;
                Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Add();
                wb.Save();
                wb.Close();
            }

            excel.DisplayAlerts = false;

            if (path.Contains(".csv"))
            {
                excel.Workbooks.Open(path, XlFileFormat.xlCSV);
            }
            else
            {
                wb = excel.Workbooks.Open(path);
            }

            ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[Sheet];

            Microsoft.Office.Interop.Excel.Range excelRange = ws.UsedRange;
            rowCount = excelRange.Rows.Count;
            colCount = excelRange.Columns.Count;
        }

        public int colCount { get; set; }
        public int rowCount { get; set; }

        public static void ConvertFromCSVtoXLSX(string csv, string xlsx)
        {
            //System.Data.DataTable dataTable = ConvertCsvToDataTable(csv); //save datatable to xlsx
            string copy = System.IO.Path.GetTempFileName();
            if (System.IO.File.Exists(copy))
            {
                System.IO.File.Delete(copy);
            }

            Application app = new Application();
            Workbook wb = app.Workbooks.Open(csv, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
            app.DisplayAlerts = false;
            wb.SaveAs(copy,
                    XlFileFormat.xlWorkbookDefault,
                    System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value,
                    false,
                    false,
                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    false,
                    false,
                  System.Reflection.Missing.Value,
                  System.Reflection.Missing.Value,
                  System.Reflection.Missing.Value);
            wb.Close();
            app.Quit();

            if (System.IO.File.Exists(xlsx))
            {
                System.IO.File.Delete(xlsx);
            }

            System.IO.File.Copy(copy, xlsx);
            System.IO.File.Delete(copy);
        }

        public static async System.Threading.Tasks.Task ConvertFromCSVtoXLSXAsync(string csv, string xlsx)
        {
            //System.Data.DataTable dataTable = ConvertCsvToDataTable(csv); //save datatable to xlsx
            string copy = System.IO.Path.GetTempFileName();
            if (System.IO.File.Exists(copy))
            {
                System.IO.File.Delete(copy);
            }

            Application app = new Application();
            Workbook wb = app.Workbooks.Open(csv, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
            app.DisplayAlerts = false;
            wb.SaveAs(copy,
                    XlFileFormat.xlWorkbookDefault,
                    System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value,
                    false,
                    false,
                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    false,
                    false,
                  System.Reflection.Missing.Value,
                  System.Reflection.Missing.Value,
                  System.Reflection.Missing.Value);
            wb.Close();
            app.Quit();

            if (System.IO.File.Exists(xlsx))
            {
                System.IO.File.Delete(xlsx);
            }

            System.IO.File.Copy(copy, xlsx);
            System.IO.File.Delete(copy);

            await System.Threading.Tasks.Task.Delay(5);
        }

        /// <summary>
        /// Reads each cell by seperating them by comma (if you have commas in the cells this could be prone to failure)
        /// </summary>
        /// <param name="csv"></param>
        /// <returns>2d string array</returns>
        public static string[][] ReadCSV(string csv)
        {
            return System.IO.File.ReadLines(csv).Select(x => x.Split(',')).ToArray();
        }

        public static Dictionary<string, Microsoft.Office.Interop.Excel.Worksheet> GetDictionairy()
        {
            return dict;
        }

        /// <summary>
        /// If the file can be opened for exclusive access it means that the file
        /// is no longer locked by another process.
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public static bool IsFileReady(string filename)
        {
            try
            {
                using System.IO.FileStream inputStream = System.IO.File.Open(filename, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.None);
                return inputStream.Length > 0;
            }
            catch (System.Exception)
            {
                return false;
            }
        }

        public static void Kill()
        {
            Process[] processes = Process.GetProcessesByName("EXCEL");
            foreach (Process p in processes)
            {
                p.Kill();
            }
        }

        /// <summary>
        /// This action saves before closing
        /// </summary>
        public void Close()
        {
            Save();
            System.GC.Collect(System.GC.MaxGeneration, System.GCCollectionMode.Forced);
            System.GC.WaitForPendingFinalizers();
            excel.Quit();
            Marshal.FinalReleaseComObject(ws);
            Marshal.FinalReleaseComObject(wb);
            Marshal.FinalReleaseComObject(excel);
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// reads the cell
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns>the object itself</returns>
        public object ReadCell(int row, int col)
        {
            if (ws.Cells[row, col] != null) return ws.Cells[row, col];
            else return null;
        }

        /// <summary>
        /// reads the cell
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns> if null or empty returns false</returns>
        public bool ReadCellBool(int row, int col)
        {
            if ((ws.Cells[row, col] as Microsoft.Office.Interop.Excel.Range).Text == null)
            {
                return false;
            }

            if (ReadCellString(row, col) == "TRUE")
            {
                return true;
            }
            else if (ReadCellInt(row, col) == 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Read cell if double will return 0 if nothing is in the cell
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">col</param>
        /// <returns>Date in cell or DateTime's max value</returns>
        public System.DateTime ReadCellDateTime(int row, int col)
        {
            System.DateTime cell = System.DateTime.MaxValue;
            bool isDateTime = System.DateTime.TryParse((ws.Cells[row, col] as Microsoft.Office.Interop.Excel.Range).Text, out cell);
            /*
            if (!isDateTime)
            {
                string cellString = ReadCellString(row, col);
                if (cellString != null)
                {
                    cell = (DateTime)Algorithms.DateManipulation.ToDate(cellString, new string[] { "mm-dd hh:mmtt yyyy" });
                }
            }*/
            return cell;
        }

        /// <summary>
        /// Read cell if double will
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">col</param>
        /// <returns>returns 0 is nothing is found</returns>
        public double ReadCellDouble(int row, int col)
        {
            double cell = 0;
            double.TryParse((ws.Cells[row, col] as Microsoft.Office.Interop.Excel.Range).Text, out cell); //this returns a bool but i dont need it
            return cell;
        }

        /// <summary>
        /// Read cell if float max will
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">col</param>
        /// <returns>returns 0 if nothing is in the cell</returns>
        public float ReadCellFloat(int row, int col)
        {
            float cell = 0;
            float.TryParse((ws.Cells[row, col] as Microsoft.Office.Interop.Excel.Range).Text, out cell); //this returns a bool but i dont need it
            return cell;
        }

        /// <summary>
        /// Read cell if int
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns>returns 0 if nothing in cell</returns>
        public int ReadCellInt(int row, int col)
        {
            int cell = 0;
            int.TryParse((ws.Cells[row, col] as Microsoft.Office.Interop.Excel.Range).Text, out cell); //this returns a bool but i dont need it
            return cell;
        }

        /// <summary>
        /// returns the entire excel sheet
        /// </summary>
        /// <param name="startRow"></param>
        /// <param name="startCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        /// <returns>2d obj arr</returns>
        public object[,] ReadCellRange()
        {
            return ws.UsedRange.Value;
        }

        /// <summary>
        /// Read the entire excel sheet faster but wont give datetimes
        /// </summary>
        /// <returns>2d object arr</returns>
        public object[,] ReadCellRangeFast()
        {
            return ws.UsedRange.Value2;
        }

        /// <summary>
        /// returns the entire excel sheet
        /// </summary>
        /// <returns>2d string arr</returns>
        public string[,] ReadCellRangeStr()
        {
            return ws.UsedRange.Text;
        }

        /// <summary>
        /// Read cell if string
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">col</param>
        /// <returns>returns the string or null if nothing found or null</returns>
        public string ReadCellString(int row, int col)
        {
            if ((string)(ws.Cells[row, col] as Microsoft.Office.Interop.Excel.Range).Text != null)
            {
                return (string)(ws.Cells[row, col] as Microsoft.Office.Interop.Excel.Range).Text;
            }
            else
            {
                return null;
            }
        }

        public string[][] ReadLines()
        {
            return System.IO.File.ReadLines(path).Select(x => x.Split(',')).ToArray();
        }

        public void Save()
        {
            wb.Save();
        }

        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }

        /// <summary>
        /// Write to cell
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">col</param>
        /// <param name="s">what you want to write</param>
        public void WriteToCell(int row, int col, string s)
        {
            (ws.Cells[row, col] as Microsoft.Office.Interop.Excel.Range).Value = s;
        }

        internal System.Data.DataTable ConvertCsvToDataTable(string filePath)
        {
            System.IO.StreamReader sr = new System.IO.StreamReader(filePath);
            string[] headers = sr.ReadLine().Split(',');
            System.Data.DataTable dataTable = new System.Data.DataTable();
            foreach (string header in headers)
            {
                dataTable.Columns.Add(header);
            }
            while (!sr.EndOfStream)
            {
                string[] rows = System.Text.RegularExpressions.Regex.Split(sr.ReadLine(), ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");
                System.Data.DataRow dr = dataTable.NewRow();
                for (int i = 0; i < headers.Length; i++)
                {
                    dr[i] = rows[i];
                }
                dataTable.Rows.Add(dr);
            }

            return dataTable;
        }

        //public static int ReleaseComObject(object o);
    }
}