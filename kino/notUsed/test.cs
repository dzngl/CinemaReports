using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using System.Data;
using System.Data.SqlClient;

namespace ReportsGenerator
{
    public class test
    {
       public List<Movie> GetMovies()
        {
            using (IDbConnection connection = new System.Data.SqlClient.SqlConnection(Helper.CnnString("kino")))
            {                    
                var output = connection.Query<Movie>($"select * from Movies").ToList();               
                return output;
                
            }

        }
        public DataTable GetMoviess()
        {
            SqlConnection conn = new SqlConnection(Helper.CnnString("kino"));            
            conn.Open();
            string query = "Select * from Movies";
            SqlCommand cmd = new SqlCommand(query, conn);
            DataTable dt = new DataTable();
            dt.Load(cmd.ExecuteReader());
            conn.Close();
            return dt;
            
        }
    }
}
