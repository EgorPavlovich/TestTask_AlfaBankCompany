using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp
{
    public class Query
    {
        public ICreateRequest Queryable { private get; set; }

        public Query(ICreateRequest Queryable)
        {
            this.Queryable = Queryable;
        }

        public void CreateQuery()
        {
            Task.Run(() =>
            {
                Queryable.CreateQuery();
            });
        }
    }
}
