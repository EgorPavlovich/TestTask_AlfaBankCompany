using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

using _Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp
{
    public class QueryByTask : Request
    {
        public override string OutputTxtfile { get; set; } = @"Output file.txt";
        public override string path { get; set; } = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), @"Output file by query.xlsx");

        public override void CreateQueryByTask()
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
                    lines[0] = phrase;
                }
                if (info.Contains(@"Номер договора"))
                {
                    String phrase = info;
                    phrase = phrase.Replace(@"Номер договора", "");
                    phrase = phrase.Trim();
                    lines[1] = phrase;
                }
                if (info.Contains(@"Счет контрагента"))
                {
                    String phrase = info;
                    phrase = phrase.Replace(@"Счет контрагента", "");
                    phrase = phrase.Trim();
                    lines[2] = phrase;
                }
                if (info.Contains(@"Адрес контрагента"))
                {
                    String phrase = info;
                    phrase = phrase.Replace(@"Адрес контрагента", "");
                    phrase = phrase.Trim();
                    lines[3] = phrase;
                }
                if (info.Contains(@"Наименование договора"))
                {
                    String phrase = info;
                    phrase = phrase.Replace(@"Наименование договора", "");
                    phrase = phrase.Trim();
                    lines[4] = phrase;
                }
            }

            // добавляем таблицу в dataSet
            dataSet.Tables.Add(dataTable);

            dataTable.Columns.Add(RegNumColumn);
            dataTable.Columns.Add(ContrNumColumn);
            dataTable.Columns.Add(InvoiceColumn);
            dataTable.Columns.Add(AddressColumn);
            dataTable.Columns.Add(NameColumn);

            DataRow row = dataTable.NewRow();
            row.ItemArray = new object[]
            {
                lines[0],
                lines[1],
                lines[2],
                lines[3],
                lines[4]
            };
            dataTable.Rows.Add(row);

            Console.Write("Регистрационный номер сделки \tНомер договора \tСчет контрагента \tАдрес контрагента \tНаименование договора");
            Console.WriteLine();
            foreach (DataRow r in dataTable.Rows)
            {
                foreach (var cell in r.ItemArray)
                    Console.Write("\t{0}", cell);
                Console.WriteLine();
            }

            WriteToExcelfile(dataSet, path);
        }

        public override void WriteToExcelfile(DataSet dataSet, string _path)
        {
            Excel excel = new Excel(_path);
            excel.GenerateFileByQuery(dataSet);
            excel.Close();
        }

    }
}
