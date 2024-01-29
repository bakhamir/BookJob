using Dapper;
using System.Data.SqlClient;
using System.Reflection;

namespace BookJob.Models
{
    public class book
    {
        public int id { get; set; }
        public string title { get; set; }
        public string author { get; set; }
        public int pages { get; set; }
        public string releaseYear { get; set; }
        public IEnumerable<book> GetBookByName(string name)
        {
            string conStr = @"Server=206-4\SQlEXPRESS;Database=testdb;Trusted_Connection=True;";
            if (!string.IsNullOrEmpty(name))
            {
                using (SqlConnection db = new SqlConnection(conStr))
                {
                    DynamicParameters parameters = new DynamicParameters();
                    parameters.Add("@name", $"{name}");
                    var res = db.Query<book>("pFindBookByName", parameters, commandType: System.Data.CommandType.StoredProcedure);
                    return res;
                }
            }
            else
            {
                return null;
            }
        }

    }

}
