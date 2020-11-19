using System.Threading.Tasks;

namespace ConsoleApp
{
    public class Query
    {
        public Request Queryable { private get; set; }

        public Query(Request Queryable)
        {
            this.Queryable = Queryable;
        }

        public void CreateQuery()
        {
            Task.Run(() =>
            {
                Queryable.CreateQueryByTask();
            });
        }
    }
}
