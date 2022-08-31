//Created by Alexander Fields https://github.com/roku674

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelCSharp
{
    public class Excel
    {
        private static Dictionary<string, Microsoft.Office.Interop.Excel.Worksheet> dict = new Dictionary<string, Worksheet>();
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
            excel.DisplayAlerts = false;

            wb = excel.Workbooks.Open(path);
            ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[Sheet];

            Microsoft.Office.Interop.Excel.Range excelRange = ws.UsedRange;
            rowCount = excelRange.Rows.Count;
            colCount = excelRange.Columns.Count;
        }

        public int colCount { get; set; }
        public int rowCount { get; set; }

        public static void ConvertFromCSVtoXLSX(string csv, string xls)
        {
            Microsoft.Office.Interop.Excel.Application xl = new Application();
            //Open Excel Workbook for conversion.
            Microsoft.Office.Interop.Excel.Workbook wb = xl.Workbooks.Open(csv);
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
            //Select The UsedRange
            Microsoft.Office.Interop.Excel.Range used = ws.UsedRange;
            //Autofit The Columns
            used.EntireColumn.AutoFit();
            //Save file as csv file
            wb.SaveAs(xls, 51);
            //Close the Workbook.
            wb.Close();
            //Quit Excel Application.
            xl.Quit();
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
        public DateTime ReadCellDateTime(int row, int col)
        {
            DateTime cell = DateTime.MaxValue;
            bool isDateTime = DateTime.TryParse((ws.Cells[row, col] as Microsoft.Office.Interop.Excel.Range).Text, out cell);
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

        //public static int ReleaseComObject(object o);
    }
}