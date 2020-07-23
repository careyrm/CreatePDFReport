using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDF_Report_Generator.DataAccess
{
    public class DataQueries
    {
        string strconnectionDev = ConfigurationManager.ConnectionStrings["DATA_DEV"].ConnectionString;
        string strconnectionProd = ConfigurationManager.ConnectionStrings["DATA_PROD"].ConnectionString;
        string strconnectionPDF_Rpt_Gen = ConfigurationManager.ConnectionStrings["PDF_REPORTS"].ConnectionString;
        string processEnvironment = ConfigurationManager.AppSettings["processEnvironment"].ToString();

        SqlConnection sqlcon = new SqlConnection();
        SqlCommand sqlcmd = new SqlCommand();
        SqlDataAdapter da = new SqlDataAdapter();
        DataTable dt = new DataTable();
        DataSet ds = new DataSet();

        public string setConnPath(string src)
        {
            if(src == "CRM")
            {
                switch (processEnvironment.ToUpper())
                {
                    case "PROD":
                        return strconnectionProd;
                    default:
                        return strconnectionDev;
                }
            }
            else
            {
                return strconnectionPDF_Rpt_Gen;
            }
        }

        public void connect(string src)
        {
            string strconnection = setConnPath(src);
            sqlcon = new SqlConnection(strconnection);
            sqlcon.Open();
        }
        public void disconnect()
        {
            if (sqlcon.State == ConnectionState.Open)
            {
                sqlcon.Close();
                sqlcon.Dispose();
            }
        }
        public DataTable ReadDataTable(string query,string src)
        {
            try
            {
                connect(src);
                sqlcmd = new SqlCommand(query, sqlcon);
                da = new SqlDataAdapter(sqlcmd);
                dt = new DataTable();
                da.Fill(dt);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                disconnect();
            }
            return dt;
        }
        public DataSet ReadDataSet(string query,string src)
        {
            try
            {
                connect(src);
                sqlcmd = new SqlCommand(query, sqlcon);
                da = new SqlDataAdapter(sqlcmd);
                ds = new DataSet();
                da.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                disconnect();
            }
            return ds;
        }
        public void QryCommand(string query,string src)
        {
            try
            {
                connect(src);
                sqlcmd = new SqlCommand(query, sqlcon);
                sqlcmd.CommandTimeout = 0;
                sqlcmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                disconnect();
            }
        }

        public void BulkSQLCopy(DataTable dt, string tablename,string src)
        {
            try
            {
                connect(src);
                using (var bulkCopy = new SqlBulkCopy(sqlcon))
                {
                    bulkCopy.BatchSize = 500;
                    bulkCopy.NotifyAfter = 1000;

                    bulkCopy.DestinationTableName = tablename;
                    bulkCopy.WriteToServer(dt);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                disconnect();
            }
        }
    }
}
