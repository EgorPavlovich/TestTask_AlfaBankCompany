using System;
using System.Data;

namespace ConsoleApp
{
    public abstract class Request
    {
        public abstract string OutputTxtfile { get; set; }
        public abstract string path { get; set; }

        public abstract void CreateQueryByTask();
        public string[] lines = new string[5];

        public DataSet dataSet = new DataSet("dataSet");
        public DataTable dataTable = new DataTable(@"Данные");

        public DataColumn RegNumColumn = new DataColumn(@"Регистрационный номер сделки", Type.GetType("System.String"));
        public DataColumn ContrNumColumn = new DataColumn(@"Номер договора", Type.GetType("System.String"));
        public DataColumn InvoiceColumn = new DataColumn(@"Счет контрагента", Type.GetType("System.String"));
        public DataColumn AddressColumn = new DataColumn(@"Адрес контрагента", Type.GetType("System.String"));
        public DataColumn NameColumn = new DataColumn(@"Наименование договора", Type.GetType("System.String"));

        public abstract void WriteToExcelfile(DataSet dataSet, string path);
    }
}
