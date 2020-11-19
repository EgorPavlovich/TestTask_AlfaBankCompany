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

        string OutputTxtfile { get; set; }
        String Registration_number_of_the_transaction { get; set; } // - «Регистрационный номер сделки»
        String Contract_number { get; set; } // - «Номер договора»
        String Counterparty_account { get; set; } // - «Счет контрагента»
        String Counteragent_address { get; set; } // - «Адрес контрагента»
        String Name_of_contract { get; set; } // - «Наименование договора»

        void WriteToExcelfile(DataTable data);
    }
}
