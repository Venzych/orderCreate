using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace orderCreateFinal.Models
{
    public class Profile
    {
        public long Id { get; set; }
        public int RecordCount { get; set; } = 0;
        public User User { get; set; }

        public long? ManagerID { get; set; } = null;
        public long? OfficeID { get; set; } = null;
        public long? PrefNumber { get; set; } = null;
        public long? PrefNumberAls { get; set; } = null;
        public long? AgentCode { get; set; } = null;
        public long? DepartID { get; set; } = null;
        public long? SellerID { get; set; } = null;
        public long? CarrierID { get; set; } = null;
        public long? MounterID { get; set; } = null;
        public long? MakerID { get; set; } = null;
        public long? WinMakerID { get; set; } = null;
        public long? GlassMakerID { get; set; } = null;
        public long? AEMakerID { get; set; } = null;
        public long? MoveConstrRuleID { get; set; } = null;
        public long? StockID { get; set; } = null;
        public long? TradeMarkID { get; set; } = null;
        public long? ContrTypeID { get; set; } = null;



        public Profile() { }

    }
}
