using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp
{
    public interface ICreateRequest
    {
        Dictionary<string, string> CreateQuery();
        //string OutputTxtfile { get; set; } //= @"Output file.txt"
        //void CreateExcelfile();

        void WriteToExcelfile(Dictionary<string, string> Query);
    }
}
