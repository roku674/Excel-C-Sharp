using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace StarportExcel
{
    internal class Excel
    {
        private static Dictionary<string, Worksheet> dict = new Dictionary<string, Worksheet>();
        private readonly string path = "";
        private _Application excel = new Microsoft.Office.Interop.Excel.Application();
        private Workbook wb;
        private Worksheet ws;

        public Excel(string path, int Sheet)
        {
            this.path = path;
            excel.DisplayAlerts = false;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
        }

        public static Dictionary<string, Worksheet> GetDictionairy()
        {
            return dict;
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

        public bool ReadCellBool(int i, int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j].Value == null)
            {
                return false;
            }

            if (ws.Cells[i, j].Value == true)
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
        /// <param name="i">row</param>
        /// <param name="j">column</param>
        /// <returns></returns>
        public double ReadCellDouble(int i, int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j].Value != null) return ws.Cells[i, j].Value;
            else return 0;
        }

        public int ReadCellInt(int i, int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j].Value != null) return (int)ws.Cells[i, j].Value;
            else return 0;
        }

        /// <summary>
        /// Read cell if string
        /// </summary>
        /// <param name="i">row</param>
        /// <param name="j">column</param>
        /// <returns></returns>
        public string ReadCellString(int i, int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j].Value != null) return ws.Cells[i, j].Value;
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
        /// <param name="i">row</param>
        /// <param name="j">column</param>
        /// <param name="s">what you want to write</param>
        public void WriteToCell(int i, int j, string s)
        {
            i++;
            j++;
            ws.Cells[i, j].Value = s;
        }

        //public static int ReleaseComObject(object o);
    }
}