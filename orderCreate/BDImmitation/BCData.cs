using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace orderCreate.BDImmitation
{
    public static class BCData
    {
        public static DataSet DataSet { get; set; }
        public static DataTable DataTable { get; set; }
        public static long UserSID { get; set; }
        public static long PullLong()
        {
            return 1;
        }
    }
}
