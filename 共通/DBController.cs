using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.DataAccess.Client;
using System.Windows.Forms;
using System.Data;

namespace MultiTaskServerInspection
{
    class DBController
    {
        private OracleConnection cnn = new OracleConnection();
        private OracleCommand cmd = new OracleCommand();
        private OracleTransaction txn = null;
        public void DBConnection(string DATABASE_USER, string DATABASE_PASSWD, string DATABASE_SOURCE)
        {
            try
            {
                string conString = "user id=" + DATABASE_USER
                    + ";password=" + DATABASE_PASSWD
                    + ";data source=" + DATABASE_SOURCE;

                cnn.ConnectionString = conString;
                cnn.Open();
            }
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message);
                MessageBox.Show((ex.Number).ToString());
            }
        }

        public void DBClose()
        {
            try
            {
                cnn.Close();
            }
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message);
                MessageBox.Show((ex.Number).ToString());
            }
        }

        public void BeginTran()
        {
            txn = cnn.BeginTransaction();
        }

        public void CommitTran()
        {
            txn.Commit();
        }

        public void Rollback()
        {
            txn.Rollback();
        }

        public DataTable ExecuteQuery(string query)
        {
            OracleCommand cmd = new OracleCommand(query, cnn);
            OracleDataAdapter oda = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            cmd.ExecuteNonQuery();
            oda.Fill(dt);
            return dt;
        }
    }
}
