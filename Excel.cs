using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelCSharp
{
    internal class Excel
    {
        public int rowCount {get; set;}
        public int colCount {get; set;}
        private static Dictionary<string, Worksheet> dict = new Dictionary<string, Worksheet>();
        private readonly string path = "";
        private _Application excel = new Application();
        private Workbook wb;
        private Worksheet ws;

        /// <summary>
        /// Excel Constructor
        /// </summary>
        /// <param name="path">Include the file extension</param>
        /// <param name="Sheet">Sheet's start at page 1 not page 0</param>
        public Excel(string path, int Sheet)
        {
            this.path = path;
            excel.DisplayAlerts = false;

            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];

            Range excelRange = ws.UsedRange;
            rowCount = excelRange.Rows.Count;
            colCount = excelRange.Columns.Count;
        }

        public static Dictionary<string, Worksheet> GetDictionairy()
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
                using (System.IO.FileStream inputStream = System.IO.File.Open(filename, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.None))
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
            GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced);
            GC.WaitForPendingFinalizers();
            excel.Quit();
            Marshal.FinalReleaseComObject(ws);
            Marshal.FinalReleaseComObject(wb);
            Marshal.FinalReleaseComObject(excel);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// reads the cell
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        public bool ReadCellBool(int row, int column)
        {
            row++;
            column++;
            if (ws.Cells[row, column].Value == null)
            {
                return false;
            }

            if (ws.Cells[row, column].Value == true || ws.Cells[row, column].Value == 1)
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
        /// <param name="column">column</param>
        /// <returns>Date in cell or new empty datetime</returns>
        public DateTime ReadCellDateTime(int row, int column)
        {
            row++;
            column++;
            if (ws.Cells[row, column].Value != null && DateTime.TryParse(ws.Cells[row, column].Value, out DateTime _)) return ws.Cells[row, column].Value;
            else return new DateTime();
        }

        /// <summary>
        /// Read cell if double will
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="column">column</param>
        /// <returns>returns max value is nothing is found</returns>
        public double ReadCellDouble(int row, int column)
        {
            row++;
            column++;
            if (ws.Cells[row, column].Value != null && double.TryParse(ws.Cells[row, column].Value, out double _)) return ws.Cells[row, column].Value;
            else return double.MaxValue;
        }

        /// <summary>
        /// Read cell if float max will
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="column">column</param>
        /// <returns>return float max if nothing is in the cell</returns>
        public float ReadCellFloat(int row, int column)
        {
            if (ws.Cells[row, column].Value != null && float.TryParse(ws.Cells[row, column].Value, out float _)) return ws.Cells[row, column].Value;
            else return float.MaxValue;
        }

        /// <summary>
        /// Read cell if int
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns>returns int max if nothing in cell</returns>
        public int ReadCellInt(int row, int column)
        {
            if (ws.Cells[row, column].Value != null && int.TryParse(ws.Cells[row, column].Value, out int _)) return (int)ws.Cells[row, column].Value;
            else return int.MaxValue;
        }

        /// <summary>
        /// Read cell if string
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="column">column</param>
        /// <returns>returns the stirng or empty if nothing found or null</returns>
        public string ReadCellString(int row, int column)
        {
            if (ws.Cells[row, column].Value != null) return ws.Cells[row, column].Value;
            else return "";
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
        /// <param name="column">column</param>
        /// <param name="s">what you want to write</param>
        public void WriteToCell(int row, int column, string s)
        {
            ws.Cells[row, column].Value = s;
        }

        //public static int ReleaseComObject(object o);
    }
}
