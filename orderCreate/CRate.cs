using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace orderCreate
{
    public class CRate
    {
        public DataTable BCData { get; set; }

        public DataTable LoadByDateInInterval(DateTime date)
        {
            DataTable dt = BCData;
            return dt;
        }

        public DataTable GetRateByDate(DateTime date, string exchangeRate)
        {
            DataTable dt = BCData;
            return dt;
        }
    }
}
