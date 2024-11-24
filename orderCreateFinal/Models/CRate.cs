using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace orderCreateFinal.Models
{
    public class CRate
    {
        public DataTable DataBase { get; set; }

        public DataTable LoadByDateInInterval(DateTime date)
        {
            // Заглушка
            DataTable dt = DataBase;
            return dt;
        }

        public double GetRateByDate(DateTime date, string exchangeRate)
        {
            DataTable dt = LoadByDateInInterval(date);
            foreach (DataRow row in dt.Rows)
            {
                if (row["Код валюты"].ToString() == exchangeRate && ((DateTime)row["Курс на дату"]).Date == date.Date)
                {
                    return Convert.ToDouble(row["Курс"]);
                }
            }
            return -1;
            /*modified

            DataTable dt = LoadByDateInInterval(date);
            double? result = dt.AsEnumerable()
            .Where(row => row.Field<string>("Код валюты") == exchangeRate 
                          && row.Field<DateTime>("Курс на дату").Date == date.Date)
            .Select(row => row.Field<double>("Курс"))
            .FirstOrDefault();
            if (result == null) return -1;
            return result.Value;

             */
        }
    }
}
