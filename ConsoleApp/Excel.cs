using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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

        public Excel(string Path)
        {
            this.Path = Path;

            excel.DisplayAlerts = false;
            // Создание объекта Excel и добавление к нему рабочей книги...
            wb = excel.Application.Workbooks.Add(true);
        }
    
        public Excel(string Path, int Sheet)
        {
            this.Path = Path;
            
            wb = excel.Workbooks.Open(Path);
            ws = wb.Worksheets[Sheet];
        }

        public void CreateNewFile(int Sheet)
        {
            this.wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            this.ws = wb.Worksheets[Sheet];
        }

        public void CreateNewSheet()
        {
            Worksheet tempSheet = wb.Worksheets.Add(After: ws);
        }

        public void GenerateFileByQuery(DataSet ds)
        {
            // Обход по циклу таблицы данных.... и добавление листов в рабочую книгу
            foreach (System.Data.DataTable dt in ds.Tables)
            {
                AddTableSheet(ref wb, ref excel, dt, true);
            }

            SaveAs(Path);
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

        public void AddTableSheet(ref Workbook wb, ref _Application excel, System.Data.DataTable dt, bool _IsHeaderIncluded)
        {

            // Переименование имя первого листа/листа по умолчанию, на название записываемой таблицы
            if (dt.TableName.Trim() != string.Empty)
            {
                Worksheet ws = (Worksheet)excel.Worksheets.get_Item(1);
                if (ws.Name.ToLower() != "sheet1")
                {
                    ws = (Worksheet)wb.Worksheets.Add();
                }
                ws.Name = dt.TableName;
            }

            int iCol = 0;
            // Добавление имён для заголовков столбцов...
            if (_IsHeaderIncluded == true)
            {
                foreach (DataColumn c in dt.Columns)
                {
                    iCol++;
                    excel.Cells[1, iCol] = c.ColumnName;

                    // (выравнивание по горизонтали) выравнивание по центру
                    (excel.Cells[1, iCol] as Range).HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    // (выравнивание по вертикали) выравнивание по центру
                    (excel.Cells[1, iCol] as Range).VerticalAlignment = XlHAlign.xlHAlignCenter;

                    // жирность текста
                    (excel.Cells[1, iCol] as Range).Font.Bold = true;
                    (excel.Cells[1, iCol] as Range).Select();

                    // бордюры
                    (excel.Cells[1, iCol] as Range).Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    (excel.Cells[1, iCol] as Range).Borders[XlBordersIndex.xlEdgeBottom].Weight = 4;

                    (excel.Cells[1, iCol] as Range).Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    (excel.Cells[1, iCol] as Range).Borders[XlBordersIndex.xlEdgeRight].Weight = 2;

                    (excel.Cells[1, iCol] as Range).Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
                    (excel.Cells[1, iCol] as Range).Borders[XlBordersIndex.xlInsideHorizontal].Weight = 2;

                    (excel.Cells[1, iCol] as Range).Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
                    (excel.Cells[1, iCol] as Range).Borders[XlBordersIndex.xlInsideVertical].Weight = 2;

                    (excel.Cells[1, iCol] as Range).Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                    (excel.Cells[1, iCol] as Range).Borders[XlBordersIndex.xlEdgeTop].Weight = 4;

                    // шрифт текста
                    (excel.Cells[1, iCol] as Range).Cells.Font.Name = "Tahoma";
                    // размер шрифта текста
                    (excel.Cells[1, iCol] as Range).Cells.Font.Size = 10.5;

                    // авто ширина текста
                    (excel.Cells[1, iCol] as Range).EntireColumn.AutoFit();

                    // высота текста
                    (excel.Cells[1, 1] as Range).Rows[1].RowHeight = 20;
                    //// авто высота текста
                    //(excel.Cells[1, iCol] as Range).EntireRow.AutoFit();
                }
            }

            // Для каждой строки данных...
            int iRow = 0;
            foreach (DataRow r in dt.Rows)
            {
                iRow++;
                // Добавление данных по ячейкам для каждой строки...
                iCol = 0;
                foreach (DataColumn c in dt.Columns)
                {
                    iCol++;
                    if (_IsHeaderIncluded == true)
                    {
                        excel.Cells[iRow + 1, iCol] = r[c.ColumnName];
                        (excel.Cells[iRow + 1, iCol] as Range).Cells.Font.Size = 11;
                        (excel.Cells[iRow + 1, iCol] as Range).HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        (excel.Cells[iRow + 1, iCol] as Range).VerticalAlignment = XlHAlign.xlHAlignCenter;
                        (excel.Cells[iRow + 1, iCol] as Range).EntireColumn.AutoFit();
                        (excel.Cells[iRow + 1, iCol] as Range).EntireRow.AutoFit();
                    }
                    else
                    {
                        excel.Cells[iRow, iCol] = r[c.ColumnName];
                        (excel.Cells[iRow + 1, iCol] as Range).Cells.Font.Size = 11;
                        (excel.Cells[iRow + 1, iCol] as Range).HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        (excel.Cells[iRow + 1, iCol] as Range).VerticalAlignment = XlHAlign.xlHAlignCenter;
                        (excel.Cells[iRow + 1, iCol] as Range).EntireColumn.AutoFit();
                        (excel.Cells[iRow + 1, iCol] as Range).EntireRow.AutoFit();
                    }
                }
            }

        }

        public void Visible(bool b)
        {
            excel.Visible = b;
        }

        public void ActivateTheWorksheet()
        {
            ws = (Worksheet)excel.ActiveSheet;
            ws.Activate();
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

        public void SaveAsTheWorkbook(string path)
        {
            // Глобально отсутствующая ссылка для объектов, которые мы не определяем...
            object missing = System.Reflection.Missing.Value;

            wb.SaveAs
            (
                Path,
                XlFileFormat.xlWorkbookNormal,
                missing,
                missing,
                false,
                false,
                XlSaveAsAccessMode.xlNoChange,
                missing,
                missing,
                missing,
                missing,
                missing
            );
        }

        public void Close()
        {
            wb.Close();
        }

        public void SaveAndClose()
        {
            wb.Close(true);
        }

        public void Shutdown()
        {
            excel.Quit();
        }

    }
}
