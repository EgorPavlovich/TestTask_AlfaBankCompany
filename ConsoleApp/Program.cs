using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp
{
    class Program
    {
        private static string OriginalRTFfile = @"testDoc.rtf";
        private static string OutputTxtfile = @"Output file.txt";

        static void Main(string[] args)
        {
            ConvertRtfToTxt(OriginalRTFfile);

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
