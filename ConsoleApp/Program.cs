using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

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

        static void Main(string[] args)
        {
            // 1.
            try
            {
                // a.
                ConvertRtfToTxt(OriginalRTFfile);

                // b.
                Query query = new Query(new QueryByTask());
                query.CreateQuery();
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
