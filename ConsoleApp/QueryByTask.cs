using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp
{
    public class QueryByTask : ICreateRequest
    {
        public Dictionary<string, string> CreateQuery()
        {
            throw new NotImplementedException();
        }

        public void WriteToExcelfile(Dictionary<string, string> Query)
        {
            throw new NotImplementedException();
        }
    }
}
