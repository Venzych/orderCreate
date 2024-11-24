using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace orderCreateFinal.Models
{
    public class TradeAgent : User
    {
        // У TradeAgent есть спец ID помимо userID
        public long TradeAgentID {  get; }
        public string Data { get; set; }

        public TradeAgent(long tradeAgentID, long id = 0, string name = "") : base(id, name)
        {
            TradeAgentID = tradeAgentID;
        }
    }
}
