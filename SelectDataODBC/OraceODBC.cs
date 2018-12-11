using Oracle.DataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SelectDataODBC
{
    /// <summary>
    /// Connects to an ODBC Data Source 
    /// </summary>
    class OracleODBC
    {
        /// <summary>
        /// Opens connection to ODBC Data Source and reads data
        /// </summary>
        /// <param name="query">SQL query</param>
        /// <returns>Returns one-dimensional array of strings</returns>
        public string[] GetData(string query, string user, string password, int noOfColumns)
        {

            // Defines connection
            string connString = "Dsn=DAR_PROD;uid= " + user + " ;pwd= " + password +" ;dbq=DAR_PROD;dba=W;apa=T;exc=F;fen=T;qto=T;frc=10;fdl=10;lob=T;rst=T;btd=F;bnf=F;bam=IfAllSuccessful;num=NLS;dpm=F;mts=T;mdi=F;csr=F;fwc=F;fbs=64000;tlo=O;mld=0;oda=F;ste=F;tsz=8192;ast=FLOAT";

            // List for data from ODBC Data Source
            List<string> rawData = new List<string>();

            // Reads the data
            using (OdbcConnection connection = new OdbcConnection(connString))
            {
                OdbcCommand command = new OdbcCommand(query, connection);
                connection.Open();

                // Execute the DataReader and access the data
                OdbcDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    for (int i = 0; i < noOfColumns; i++)
                        rawData.Add(reader[i].ToString());
                }

                string[] data = rawData.ToArray();

                // Done reading
                reader.Close();

                return data;
            }
        }
    }
}
