using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.DataAccess.Client;

namespace MultiTaskServerInspection
{
    class DBRS
    {
        public DataTable TablespaceStatement(string UserName, string Passwd, string InstanceName)
        {
            DBController controller = new DBController();
            StringBuilder query = new StringBuilder(128);
            DataTable dt = null;
            try
            {
                controller.DBConnection(UserName, Passwd, InstanceName);
                query.Append("select ");
                query.Append("x.TABLESPACE_NAME,");
                query.Append("x.free_GByte,");
                query.Append("trunc(sum(u.bytes / power(1024, 3)), 3) as Total_Gbyte ");
                query.Append("from ");
                query.Append("(select ");
                query.Append("t.TABLESPACE_NAME,");
                query.Append("trunc(sum(t.bytes / power(1024, 3)), 2) as free_GByte ");
                query.Append("from dba_free_space t ");
                query.Append("group by t.TABLESPACE_NAME");
                query.Append(")x ");
                query.Append("left outer join dba_data_files u on x.TABLESPACE_NAME = u.TABLESPACE_NAME ");
                query.Append("where x.TABLESPACE_NAME in ('TSK', 'SYSTEM', 'SYSAUX') ");
                query.Append("group by ");
                query.Append("x.TABLESPACE_NAME,");
                query.Append("x.free_GByte");
                dt = controller.ExecuteQuery(query.ToString());
                return dt;
            }
            catch (OracleException ex)
            {
                throw ex;
            }
            finally
            {
                dt = null;
                if(controller != null)
                {
                    controller.DBClose();
                }
                controller = null;
            }
        }
    }
}
