using orderCreateFinal.Models;
using System.Data;
using System.Runtime;

namespace orderCreateFinal
{
    internal class Program
    {
        static void Main(string[] args)
        {
            DataSet dataSet = new DataSet();
            DataSetter dataSetter = new DataSetter(dataSet);

            CRate cRate = new CRate();
            cRate.DataBase = dataSetter.GetDataTableByName("CRate");

            Codificator codif = new Codificator();
            codif.DataBase = dataSetter.GetDataTableByName("Codificator");

            UserDB userDB = new UserDB();
            userDB.DataBase = dataSetter.GetDataTableByName("UserDB");

            Props props = new Props();
            props.DataBase = dataSetter.GetDataTableByName("Props");

            FcdOrder fcdOrder = new FcdOrder();
            fcdOrder.DataBase = dataSetter.GetDataTableByName("FcdOrder");


            // Курс на текущую дату
            double Exchange = cRate.GetRateByDate(DateTime.Now, "TRATE_EUR");

            // Условия оплаты "по стандартной скидке"
            long payCondID;
            long? row = (long)dataSetter.GetValueByColumnCondition(dataSetter.GetDataTableByName("PAYMENTCOND"), "Условия оплаты", "STDSC", "Идентификатор условия оплаты");
            if (row != null) 
            {
                payCondID = row.Value;
            }


            Order order = new Order(dataSet);

            long id = (long)dataSetter.GiveData("UserID");
            string name = (string)dataSetter.GiveData("name");
            User user = new User(id, name);
            order.UserID = user.Id;

            Profile profile = new Profile();
            if ((long)dataSetter.GiveData("ProfileID") != 0) profile.Id = (long)dataSetter.GiveData("ProfileID");
            else profile.Id = (long)dataSetter.GiveData("ByUserID");
            order.ProfileID = profile.Id;

            long ManagerID = 0;
            long OfficeID = 0;
            long DepartID = 0;
            long SellerID = 0;
            long CarrierID = 0;
            long MounterID = 0;
            long MakerID = 0;
            long WinMakerID = 0;
            long GlassMakerID = 0;
            long AEMakerID = 0;
            long MoveConstrRuleID = 0;
            long StockID = 0;
            long TradeMarkID = 0;
            long ContrTypeID = 0;
            if (profile.RecordCount > 0)
            {
                ManagerID = profile.ManagerID ?? 0;
                OfficeID = profile.OfficeID ?? 0;
                DepartID = profile.DepartID ?? 0;
                SellerID = profile.SellerID ?? 0;
                CarrierID = profile.CarrierID ?? 0;
                MounterID = profile.MounterID ?? 0;
                MakerID = profile.MakerID ?? 0;
                WinMakerID = profile.WinMakerID ?? 0;
                GlassMakerID = profile.GlassMakerID ?? 0;
                AEMakerID = profile.AEMakerID ?? 0;
                MoveConstrRuleID = profile.MoveConstrRuleID ?? 0;
                StockID = profile.StockID ?? 0;
                TradeMarkID = profile.TradeMarkID ?? 0;
                ContrTypeID = profile.ContrTypeID ?? 0;
            }


            long otmp;
            string onum;
            string onumA;
            bool isChangeDF = false;
            isChangeDF = false;
            DataTable BCData = dataSetter.GetDataTableByName("Office");
            DataRow? rowBCData = dataSetter.GetRowByColumnValue(BCData, "Ofice", $"{OfficeID}");
            otmp = Convert.ToInt64(rowBCData["OfficeSID"]);
            onum = rowBCData["OfficeNum"].ToString();
            onumA = rowBCData["OfficeNumAlias"].ToString();

            rowBCData["OfficeSID"] = OfficeID;

            if (profile.RecordCount > 0)
            {
                rowBCData["OfficeNum"] = profile?.PrefNumber.Value.ToString().Trim();
                rowBCData["OfficeNumAlias"] = profile?.PrefNumberAls.Value.ToString().Trim();
            }
            
            profile.OfficeID = otmp;


            // Задать торгового агента
            long TradeAgentID = 0;
            if (TradeAgentID != 0)
            {
                order.TradeAgentID = TradeAgentID;
                TradeAgent tradeAgent = new TradeAgent(tradeAgentID: 3726343);
                BCData = dataSetter.GetDataTableByName("TradeAgent");
                rowBCData = dataSetter.GetRowByColumnValue(BCData, "TradeAgentID", $"{tradeAgent.TradeAgentID}");
                tradeAgent.Data = rowBCData["Data"].ToString();

                order.TradeAgentID = tradeAgent.TradeAgentID;
            }

            // Задать номер дисконтной карты
            if (!string.IsNullOrWhiteSpace(order.DCardNum))
            {
                dataSetter.ToData(order.DCardNum);
            }

            // Проверка, если заказ по замеру платного ремонта и менеджер заказа принадлежит к группе "Замерщики"
            if (order.MeterReq.["ContactKindKey"].ToString().Trim() == "RP" &&
                profile.GetUserGroupKeyByManagerID(ManagerID) == "METER" &&
                MeterReq["TradeMarkID"].ToString() == DBProcLib.GetIDByKey(BCData, "TRADEMARK", "MO"))
            {
                OfficeID = 125;
            }
        }
                
    }
}
