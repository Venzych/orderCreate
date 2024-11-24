using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Text;
using System.Threading.Tasks;

namespace orderCreate.BDImmitation
{
    public class DBProcLib
    {
        public DataTable DataTable { get; set; }

        public static long GetIDByKey(DataSet dataSet, string tableName, string key)
        {
            DataTable paymentCondTable = dataSet.Tables[$"{tableName}"];
            DataRow[] rows = paymentCondTable.Select($"{tableName} = '{key}'");
            long result;
            if (long.TryParse(rows[0]["PAYMENTCOND"].ToString(), out result)) return result;
            return -1;
        }
    }
}
