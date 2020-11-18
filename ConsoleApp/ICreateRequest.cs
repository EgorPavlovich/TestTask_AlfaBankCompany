using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp
{
    public interface ICreateRequest
    {
        Dictionary<string, string> CreateQuery();

        String Registration_number_of_the_transaction { get; set; }
        String Contract_number { get; set; }
        String Counterparty_account { get; set; }
        String Counteragent_address { get; set; }
        String Name_of_contract { get; set; }

        void WriteToExcelfile(DataTable data);
    }
}
