using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp
{
    public class Excel
    {
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public string Path { get; set; } = string.Empty;

        public Excel() { }
    
        public Excel(string Path, int Sheet)
        {
            this.Path = Path;
            
            wb = excel.Workbooks.Open(Path);
            ws = wb.Worksheets[Sheet];
        }

        public void CreateNewFile()
        {
            this.wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            this.ws = wb.Worksheets[1];
        }

        public void CreateNewSheet()
        {
            Worksheet tempSheet = wb.Worksheets.Add(After: ws);
        }

        public string ReadCell(int i, int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j].Value2 != null)
                return ws.Cells[i, j].Value2;
            else
                return string.Empty;
        }

        public void WriteToCell(int i, int j, string s)
        {
            i++;
            j++;
            ws.Cells[i, j].Value2 = s;
        }

        public string[,] ReadRange(int starti, int starty, int endi, int endy)
        {
            Range range = (Range) ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            object[,] holder = range.Value2;
            string[,] result = new string[endi - starti, endy - starty];
            for (int p = 1; p <= endi - starti; p++)
            {
                for (int q = 1; q <= endy - starty; q++)
                {
                    result[p - 1, q - 1] = holder[p, q].ToString();
                }
            }

            return result;
        }

        public void WriteRange(int starti, int starty, int endi, int endy, string[,] text)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            range.Value2 = text;
        }

        public void SelectWorksheet(int SheetNumber)
        {
            this.ws = wb.Worksheets[SheetNumber];
        }

        public void DeleteWorksheet(int SheetNumber)
        {
            wb.Worksheets[SheetNumber].Delete();
        }

        public void ProtectSheet()
        {
            ws.Protect();
        }

        public void ProtectSheet(string Password)
        {
            ws.Protect(Password);
        }

        public void UnprotectSheet()
        {
            ws.Unprotect();
        }

        public void UnprotectSheet(string Password)
        {
            ws.Unprotect(Password);
        }

        public void Save()
        {
            wb.Save();
        }

        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }

        public void Close()
        {
            wb.Close();
        }

    }
}
