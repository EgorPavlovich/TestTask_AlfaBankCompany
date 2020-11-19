using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp
{
    /// <summary>
    /// Тестовое задание.
    /// 1. Создать консольное приложение, которое:
    /// </summary>
    /// <remarks>
    /// a. Из файла testDoc.rtf достает следующую информацию:
    /// <list type="table">
    /// <item>
    /// <description>- Регистрационный номер сделки</description>
    /// </item>
    /// <item>
    /// <description>- Номер договора</description>
    /// </item>  
    /// <item>
    /// <description>- Счет контрагента</description>
    /// </item>
    /// <item>
    /// <description>- Адрес контрагента</description>
    /// </item>
    /// <item>
    /// <description>- Наименование договора</description>
    /// </item>  
    /// </list>
    /// b. Создает Excel-файл, со столбцами:
    /// <list type="table">
    /// <item>
    /// <description>- «Регистрационный номер сделки»</description>
    /// </item>
    /// <item>
    /// <description>- «Номер договора»</description>
    /// </item>  
    /// <item>
    /// <description>- «Счет контрагента»</description>
    /// </item>
    /// <item>
    /// <description>- «Адрес контрагента»</description>
    /// </item>
    /// <item>
    /// <description>- «Наименование договора»</description>
    /// </item>  
    /// </list>
    /// c. Заполняет столбцы в п.2 данными из п.1.
    /// </remarks>
    /// <summary>
    /// </summary> 
    class Program
    {
        private static string OriginalRTFfile = @"testDoc.rtf";
        private static string OutputTxtfile = @"Output file.txt";
        private static string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), @"Output file by query.xlsx");

        static void Main(string[] args)
        {
            //ConvertRtfToTxt(OriginalRTFfile);
            try
            {
                DataSet dataSet = new DataSet("dataSet");
                DataTable dataTable = new DataTable(@"Данные");
                // добавляем таблицу в dataset
                dataSet.Tables.Add(dataTable);

                // создаем столбцы для таблицы Данные
                DataColumn RegNumColumn = new DataColumn(@"Регистрационный номер сделки", Type.GetType("System.String"));
                DataColumn ContrNumColumn = new DataColumn(@"Номер договора", Type.GetType("System.String"));
                DataColumn InvoiceColumn = new DataColumn(@"Счет контрагента", Type.GetType("System.String"));
                DataColumn AddressColumn = new DataColumn(@"Адрес контрагента", Type.GetType("System.String"));
                DataColumn NameColumn = new DataColumn(@"Наименование договора", Type.GetType("System.String"));

                dataTable.Columns.Add(RegNumColumn);
                dataTable.Columns.Add(ContrNumColumn);
                dataTable.Columns.Add(InvoiceColumn);
                dataTable.Columns.Add(AddressColumn);
                dataTable.Columns.Add(NameColumn);

                DataRow row = dataTable.NewRow();
                row.ItemArray = new object[]
                    { "211112/211122/211234",
                        "12992617381936",
                        "BY11PJCB30119349931000000933",
                        "ТОКИО, УЛ.КИМЧЕНЫРОВИЧА, Д.11",
                        "Тестовый договор на проводку в ДЖАПАН ТАБАСКО"
                    };
                dataTable.Rows.Add(row); // добавляем первую строку
                //dataTable.Rows.Add(new object[] { null, , 300 }); // добавляем вторую строку


                Console.Write("Регистрационный номер сделки \tНомер договора \tСчет контрагента \tАдрес контрагента \tНаименование договора");
                Console.WriteLine();
                foreach (DataRow r in dataTable.Rows)
                {
                    foreach (var cell in r.ItemArray)
                        Console.Write("\t{0}", cell);
                    Console.WriteLine();
                }

                Excel excel = new Excel(path);
                excel.GenerateFileByQuery(dataSet);
                excel.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            Console.ReadLine();
        }

        public static void ConvertRtfToTxt(string path_)
        {
            string path = path_;
            System.Windows.Forms.RichTextBox rtBox = new System.Windows.Forms.RichTextBox();
            string rtfText = System.IO.File.ReadAllText(path);
            rtBox.Rtf = rtfText;
            string plainText = rtBox.Text;

            using (StreamWriter file = new StreamWriter(OutputTxtfile, true, System.Text.Encoding.Default))
            {
                file.WriteLine(plainText);
            }

            // читаем все данные из файла
            string[] lines = System.IO.File.ReadAllLines(OutputTxtfile, Encoding.Default);
            // преобразуем в список
            var list = new List<string>(lines);
            // получаем только уникальные элементы
            var uniqueStrings = list.Distinct();
            // записываем их обратно                      
            System.IO.File.WriteAllLines(OutputTxtfile, uniqueStrings, System.Text.Encoding.Default);
        }


    }
}
