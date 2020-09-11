using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using _Excel = Microsoft.Office.Interop.Excel;

namespace Schedule_Calculator_Pro
{
    public class Excel
    {
        string path = "";
        public _Application excel = new _Excel.Application();
        public Workbook wb;
        public Worksheet ws;
        public Excel(string path)
        {
            Kill(path);
            this.path = path;
            excel.SheetsInNewWorkbook = 1;
            wb = excel.Workbooks.Add(1);
            ws = wb.Worksheets[1];
        }
        public Excel(string path, int sheet)
        {
            Kill(path);
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }
        public void close()
        {
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            Kill(path);
        }

        public string ReadCell(int i, int j)
        {
            return Convert.ToString(ws.Cells[i + 1, j + 1].Value2);
        }
        public void WriteToCell(int i, int j, string s)
        {
            ws.Cells[i + 1, j + 1].Value2 = s;
        }
        public void Save()
        {
            wb.Save();
        }
        public void SaveAs()
        {
            if (File.Exists(path))
                File.Delete(path);
            wb.SaveAs(path);
        }
        public bool BReadCell(int i, int j)
        {
            i++; j++;
            if (ws.Cells[i, j].Value2 != null)
                return true;
            else
                return false;
        }

        public static void Kill(string excelFileName) // убиваем процесс по имени файла
        {
            var processes = from p in Process.GetProcessesByName("EXCEL") select p;

            foreach (var process in processes)
                if (process.MainWindowTitle == "Microsoft Excel - " + excelFileName)
                    process.Kill();
        }
    }
}
