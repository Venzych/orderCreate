using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace orderCreateFinal.Models
{
    public class Order
    {
        public long MetReqNum { get; set; }
        public DataSet MeterReq { get; set; }
        public long UserID { get; set; } = 0;
        public long ProfileID { get; set; } = 0;
        public long TradeAgentID { get; set; } = 0;
        public string DCardNum { get; set; } = "";



        public Order (DataSet meterReq)
        {
            MeterReq = meterReq;
            
        }
        public Order (DataSet meterReq, long userID, long profileID, long tradeAgentID, string dCardNum)
        {
            MeterReq = meterReq;
            UserID = userID;
            ProfileID = profileID;
            TradeAgentID = tradeAgentID;
            DCardNum = dCardNum;
        }




    }
}
