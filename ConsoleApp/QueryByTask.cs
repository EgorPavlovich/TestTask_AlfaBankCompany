using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

using _Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp
{
    public class QueryByTask : Excel, ICreateRequest
    {
        public string Registration_number_of_the_transaction { get; set; } = String.Empty;
        public string Contract_number { get; set; } = String.Empty;
        public string Counterparty_account { get; set; } = String.Empty;
        public string Counteragent_address { get; set; } = String.Empty;
        public string Name_of_contract { get; set; } = String.Empty;

        public Dictionary<string, string> CreateQuery()
        {
            List<string> infoList = new List<string>();
            FileStream file_ = new FileStream(OutputTxtfile, FileMode.Open);
            using (StreamReader readFile = new StreamReader(file_, System.Text.Encoding.Default))
            {
                while (!readFile.EndOfStream)
                {
                    infoList.Add(readFile.ReadLine());
                }
            }

            foreach (var info in infoList)
            {
                if (info.Contains(@"Регистрационный номер сделки"))
                {
                    String phrase = info;
                    phrase = phrase.Replace(@"Регистрационный номер сделки", "");
                    phrase = phrase.Trim();
                    Registration_number_of_the_transaction = phrase;
                }
                if (info.Contains(@"Номер договора"))
                {
                    String phrase = info;
                    phrase = phrase.Replace(@"Номер договора", "");
                    phrase = phrase.Trim();
                    Contract_number = phrase;
                }
                if (info.Contains(@"Счет контрагента"))
                {
                    String phrase = info;
                    phrase = phrase.Replace(@"Счет контрагента", "");
                    phrase = phrase.Trim();
                    Counterparty_account = phrase;
                }
                if (info.Contains(@"Адрес контрагента"))
                {
                    String phrase = info;
                    phrase = phrase.Replace(@"Адрес контрагента", "");
                    phrase = phrase.Trim();
                    Counteragent_address = phrase;
                }
                if (info.Contains(@"Наименование договора"))
                {
                    String phrase = info;
                    phrase = phrase.Replace(@"Наименование договора", "");
                    phrase = phrase.Trim();
                    Name_of_contract = phrase;
                }
            }

            Dictionary<string, string> columns = new Dictionary<string, string>(5);
            columns.Add(@"Регистрационный номер сделки", $@"{Registration_number_of_the_transaction}");
            columns.Add(@"Номер договора", $@"{Contract_number}");
            columns.Add(@"Счет контрагента", $@"{Counterparty_account}");
            columns.Add(@"Адрес контрагента", $@"{Counteragent_address}");
            columns.Add(@"Наименование договора", $@"{Name_of_contract}");

            foreach (KeyValuePair<string, string> keyValue in columns)
            {
                Console.WriteLine(keyValue.Key + ": " + keyValue.Value);
            }
            Console.WriteLine();

            return columns;
        }

        public void WriteToExcelfile(DataTable data)
        {
            try
            {
                _Excel.Application EoXL;
                _Excel._Workbook EoWB;
                _Excel._Worksheet EoSheet;
                _Excel.Range excelRange;
                EoXL = new _Excel.Application();
                EoXL.Visible = false;
                EoWB = EoXL.Workbooks.Add(Type.Missing);

                int TabRows = 1;

                EoSheet = (_Excel.Worksheet)EoWB.Worksheets.get_Item(1);//ссылка на лист excel
                EoSheet.Name = "Отчет о кодах возвратных накладных";
                EoSheet.PageSetup.Orientation = _Excel.XlPageOrientation.xlLandscape;

                int row = data.Rows.Count;
                int col = data.Columns.Count;


                EoSheet.Cells[1, 1] = "Префиксы возвратных накладных и счетов фактур подразделений";
                EoSheet.Cells[1, 1].VerticalAlignment = _Excel.XlVAlign.xlVAlignCenter;
                EoSheet.Cells[1, 1].Font.Bold = true;
                EoSheet.Cells[1, 1].Font.Size = 16;

                // передаем первую таблицу, заполняем ее в памяти и передаем целиком
                object[,] dataExport = new object[row, col];

                for (int i = 0; i < row; i++)
                {
                    for (int j = 0; j < col; j++)
                    {
                        dataExport[i, j] = data.Rows[i][j];
                    }

                }

                excelRange = EoSheet.Range[EoSheet.Cells[2 + TabRows, 1], EoSheet.Cells[row + 1 + TabRows, col]];
                excelRange.set_Value(_Excel.XlRangeValueDataType.xlRangeValueDefault, dataExport);
                excelRange.Borders.ColorIndex = 0;

                //этот кусок в качестве примера указания типа данных в ячейках
                // excelRange = EoSheet.Range[EoSheet.Cells[2 + TabRows, 8], EoSheet.Cells[row + 1 + TabRows, 10]];
                // excelRange.NumberFormat = "#,##0.00";

                // формируем заголовок
                ArrayList displayColumnExsel = new ArrayList();

                foreach (DataColumn c in data.Columns)
                {

                    displayColumnExsel.Add(c.ColumnName);
                }

                object[] dataExportH = new object[col];
                for (int i = 0; i < col; i++)
                    dataExportH[i] = displayColumnExsel[i];

                excelRange = EoSheet.Range[EoSheet.Cells[1 + TabRows, 1], EoSheet.Cells[1 + TabRows, col]];
                excelRange.set_Value(_Excel.XlRangeValueDataType.xlRangeValueDefault, dataExportH);
                excelRange.Font.Bold = true;
                excelRange.WrapText = true;
                excelRange.Borders.ColorIndex = 0;
                excelRange.VerticalAlignment = _Excel.XlVAlign.xlVAlignCenter;
                excelRange.HorizontalAlignment = _Excel.XlHAlign.xlHAlignCenter;

                EoXL.Visible = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message, "Ошибка метода переноса таблиц");
            }
        }
    }
}
