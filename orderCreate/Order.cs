using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using orderCreate.BDImmitation;

namespace orderCreate
{
    public class Order
    {
        public long UserID { get; set; }
        public long ProfileID { get; set; }
        public long TradeAgentID { get; set; }
        public string DCardNum { get; set; }

        public string StrReturnString { get; set; }
        public long LngOrderID { get; set; }
        public long LngPayCondID { get; set; }
        public double DblExchange { get; set; }
        public string StrCommandNumber { get; set; }
        public long LngCurrencyID { get; set; }
        public string StrMetReqNum { get; set; }
        public long LngMeterID { get; set; }
        public long LngDocFormID { get; set; }
        public long LngManagerID { get; set; }
        public long LngOfficeID { get; set; }
        public long LngDepartID { get; set; }
        public long LngSellerID { get; set; }
        public long LngCarrierID { get; set; }
        public long LngMounterID { get; set; }
        public long LngMakerID { get; set; }
        public long LngWinMakerID { get; set; }
        public long LngGlassMakerID { get; set; }
        public long LngAEMakerID { get; set; }
        public long LngTradeMarkID { get; set; }
        public long Rc { get; set; }
        public DataTable RstOrder { get; set; }
        public DataTable RstProps { get; set; }
        public DataTable RstProfile { get; set; }
        public OrdMileSt ObjOrdMileSt { get; set; }
        public DataTable RstOrdMilest { get; set; }
        public CCurrency ObjCurrency { get; set; }
        public CRate ObjCRate { get; set; }
        public Codificator ObjCodif { get; set; }
        public DataTable RstCRate { get; set; }
        public Dealer ObjDealerData { get; set; }
        public DataTable RstDealerData { get; set; }
        public Metering ObjMetReq { get; set; }
        public List<DateTime> ColMSDates { get; set; }
        public DataTable RstVendor { get; set; }
        public DataTable RstMaker { get; set; }
        public User ObjUser { get; set; }
        public OrderProp ObjProp { get; set; }
        public fcdOrder ObjFcdOrder { get; set; }
        public long LngMoveConstrRuleID { get; set; }
        public long LngStockID { get; set; }
        public long LngShipWayID { get; set; }
        public long LngContrTypeID { get; set; }
        public long LngContrTypeID1 { get; set; }

        public Order()
        {
            ObjProp = new OrderProp();
            ObjOrdMileSt = new OrdMileSt();
            ObjOrdMileSt.BCData = BCData.DataTable;
            ObjCRate = new CRate();
            ObjCodif = new Codificator();
            ObjUser = new User();
            ObjFcdOrder = new fcdOrder();

            StrReturnString = "";

            // передать ссылку на общие данные
            ObjCRate.BCData = BCData.DataTable;
            ObjCodif.BCData = BCData.DataTable;
            ObjUser.BCData = BCData.DataTable;
            ObjProp.BCData = BCData.DataTable;
            ObjFcdOrder.BCData = BCData.DataTable;
        }

        // Получить курсы на текущую дату
        public DataTable GetCourses()
        {
            DataTable rstCRate = ObjCRate.LoadByDateInInterval(DateTime.Now);
            string strCommandNumber = (RstCRate.Rows[0]["strCommandNumber"]?.ToString() ?? "").Trim();
            string exchangeRate = "TRATE_EUR";
            DataTable dblExchange = ObjCRate.GetRateByDate(DateTime.Now, exchangeRate);
            return dblExchange;
        }


        // Условия оплаты = "по стандартной скидке"
        public long GetPrice(DataSet ObjBCData)
        {
            LngPayCondID = DBProcLib.GetIDByKey(ObjBCData, "PAYMENTCOND", "STDSC");
            return LngPayCondID;
        }

        // Проверка/заполнение данных
        public void Checkout()
        {
            if (UserID == 0) UserID = BCData.PullLong();

        }
    }
}
