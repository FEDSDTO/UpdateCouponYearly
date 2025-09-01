using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace UpdateCouponYearly
{
    class UpdateCoupon
    {
        public static DataTable CouponDeferred_Get(DateTime _sDate, DateTime _eDate)
        {
            DB_Connection _db = new DB_Connection();
            List<SqlParameter> _Parameter = new List<SqlParameter>();
            string _SQL = @"Select A.Year,A.MallId,A.GId,M.PosMallId+GiftsNo GiftsNo,UR.ExchangeStart,UR.ExchangeEnd,UR.UsedStart,UR.UsedEnd,
                                    Round(Sum(NPrice)/1.05,0) NPrice,Round(Sum(UPrice)/1.05,0) UPrice,Round(Sum(NPrice)/1.05,0) - Round(Sum(UPrice)/1.05,0) TotalPrice	
                                    From (    
                                    		Select G.Year,C.MallId,GId,G.GiftsNo,Count(GId)*G.Denomination NPrice,0 UPrice,UsedStart,UsedEnd 
                                    		 From Coupon C
                                    		 join Gifts G 
                                    		 on C.GId = G.Id  and G.Type = 'C' And C.Status != 'F' And C.CreateOn < @EDate And G.SAPType <> 'B'
                                    		 Group by C.MallId,GId,GiftsNo,UsedStart,UsedEnd,Denomination,G.Year
                                            union all
                                            Select G.Year,C.MallId,C.GId,G.GiftsNo,0 NPrice,Count(C.GId)*G.Denomination UPrice,C.UsedStart,C.UsedEnd
                                             From Coupon C
                                             join Gifts G  on C.GId = G.Id  and Status = 'U' And G.Type = 'C' And C.CreateOn < @EDate And G.SAPType <> 'B'  And Len(C.MemberId) < 10 And MemberId <> ''
                                             Group by C.MallId,C.GId,GiftsNo,C.UsedStart,C.UsedEnd,Denomination,G.Year
                                            )A
                                    join Mall M on A.MallId = M.MallId
                                    join UsedRule UR on A.GId = UR.GId And A.MallId = UR.MallId And UR.IsUse = 1 AND UR.UsedEnd >= @SDate ANd UR.UsedEnd < @EDate AND UR.UsedStart < @EDate
                                    Group by A.MallId,A.GId,GiftsNo,PosMallId,A.UsedStart,A.UsedEnd,UR.ExchangeStart,UR.ExchangeEnd,UR.UsedStart,UR.UsedEnd,A.Year
                                                ";
            _Parameter.Add(new SqlParameter("SDate", _sDate));
            _Parameter.Add(new SqlParameter("EDate", _eDate));

            return _db.GetDataTable(_SQL, _Parameter);
        }
        /// <summary>
        /// 年初年度遞延報表
        /// </summary>

        public static void Deferred_SubmitYearly()
        {
            CommonUtility commonUtility = new CommonUtility();
            try
            {
                DateTime _before = DateTime.Now.AddYears(-1);
                DateTime _sDate = new DateTime(_before.Year, 1, 1);//去年初
                DateTime _eDate = new DateTime(DateTime.Now.Year, 1, 1);//今年初
                if (ConfigurationManager.AppSettings["DeBUG"] == "Y")
                {
                    DateTime.TryParseExact(ConfigurationManager.AppSettings["startDate"], "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateS);
                    _sDate = dateS;
                    DateTime.TryParseExact(ConfigurationManager.AppSettings["endDate"], "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateE);
                    _eDate = dateE;
                }
                DataTable _dt = CouponDeferred_Get(_sDate, _eDate);
                List<string> _listMall = new List<string>() { "32", "34", "37", "40", "42", "48", "50", "51", "52", "53", "54", "55", "72" };
                //List<string> _listMall = new List<string>() { "72" };
                foreach (var _mall in _listMall)
                {
                    long _TotalPrice = 0;
                    long _UsedPrice = 0;
                    XLWorkbook workbook = new XLWorkbook();
                    var sheet = workbook.Worksheets.Add("年度兌回");
                    int _num = 6; int nPrice = 0; int uPrice = 0; int totalPrice = 0;
                    DataRow[] _drTemp = _dt.Select("MallId=" + _mall);
                    sheet.Cell(1, 1).Value = "查詢日期：" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    sheet.Cell(2, 1).Value = "查詢人員：系統排程匯出";
                    sheet.Cell(3, 1).Value = "查詢公司：" + _mall;
                    sheet.Cell(4, 1).Value = $"兌回起訖：{_sDate.ToString("yyyy-MM-dd")}至{_eDate.ToString("yyyy-MM-dd")}";
                    sheet.Cell(5, 1).Value = "券年度";
                    sheet.Cell(5, 2).Value = "券代號";
                    sheet.Cell(5, 3).Value = "兌出起始";
                    sheet.Cell(5, 4).Value = "兌出結束";
                    sheet.Cell(5, 5).Value = "兌回起始";
                    sheet.Cell(5, 6).Value = "兌回結束";
                    sheet.Cell(5, 7).Value = "兌出(除稅)";
                    sheet.Cell(5, 8).Value = "兌回(除稅)";
                    sheet.Cell(5, 9).Value = "兌出-兌回(除稅)";

                    if (_drTemp.Count() > 0)
                    {

                        foreach (var item in _drTemp)
                        {
                            sheet.Cell(_num, 1).Value = item["Year"].ToString();
                            sheet.Cell(_num, 2).Value = item["GiftsNo"].ToString();
                            if (!string.IsNullOrEmpty(item["ExchangeStart"].ToString()))
                            {
                                DateTime.TryParse(string.Format("{0:yyyy/MM/dd}", item["ExchangeStart"]), out DateTime ExchangeStart); //string 轉成datetime
                                sheet.Cell(_num, 3).Value = ExchangeStart;
                            }
                            else
                            {
                                sheet.Cell(_num, 3).Value = item["ExchangeStart"].ToString();
                            }

                            if (!string.IsNullOrEmpty(item["ExchangeEnd"].ToString()))
                            {

                                DateTime.TryParse(string.Format("{0:yyyy/MM/dd}", item["ExchangeEnd"]), out DateTime ExchangeEnd); //string 轉成datetime
                                sheet.Cell(_num, 4).Value = ExchangeEnd;
                            }
                            else
                            {
                                sheet.Cell(_num, 4).Value = item["ExchangeEnd"].ToString();
                            }

                            if (!string.IsNullOrEmpty(item["UsedStart"].ToString()))
                            {
                                //轉datetime 兌回率計算才能找出最大最小值
                                DateTime.TryParse(string.Format("{0:yyyy/MM/dd}", item["UsedStart"]), out DateTime UsedStart); //string 轉成datetime
                                sheet.Cell(_num, 5).Value = UsedStart;
                            }
                            else
                            {
                                sheet.Cell(_num, 5).Value = item["UsedStart"].ToString();
                            }

                            if (!string.IsNullOrEmpty(item["UsedEnd"].ToString()))
                            {
                                //轉datetime 兌回率計算才能找出最大最小值
                                DateTime.TryParse(string.Format("{0:yyyy/MM/dd}", item["UsedEnd"]), out DateTime UsedEnd); //string 轉成datetime
                                sheet.Cell(_num, 6).Value = UsedEnd;
                            }
                            else
                            {
                                sheet.Cell(_num, 6).Value = item["UsedEnd"].ToString();
                            }
                            //sheet.Cell(_num, 3).Value = string.Format("{0:yyyy/MM/dd}", item["ExchangeStart"]);
                            //sheet.Cell(_num, 4).Value = string.Format("{0:yyyy/MM/dd}", item["ExchangeEnd"]);
                            sheet.Cell(_num, 7).Value = Convert.ToInt32(item["NPrice"]);
                            sheet.Cell(_num, 8).Value = Convert.ToInt32(item["UPrice"]);
                            sheet.Cell(_num, 9).Value = Convert.ToInt32(item["TotalPrice"]);
                            sheet.Range($"A{_num}:I{_num}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right; //靠右
                            _num++;
                        }
                        sheet.ColumnWidth += 8;

                    }
                    //年度遞延需計算兌回率
                    var sheet2 = workbook.Worksheets.Add("兌回率計算");
                    sheet2.Cell(1, 1).Value = "查詢日期：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    sheet2.Cell(2, 1).Value = "查詢人員：系統排程匯出";
                    sheet2.Cell(3, 1).Value = "查詢公司：" + _mall;
                    sheet2.Cell(4, 1).FormulaA1 = "=年度兌回!A4";
                    sheet2.Cell(5, 1).Value = "兌出期間";
                    sheet2.Cell(5, 2).Value = "兌回起日";
                    sheet2.Cell(5, 3).Value = "兌回迄日";
                    sheet2.Cell(5, 4).Value = "兌出";
                    sheet2.Cell(5, 5).Value = "兌回";
                    sheet2.Cell(5, 6).Value = "兌回率";
                    sheet2.Cell(6, 1).FormulaA1 = "=TEXT(MIN(年度兌回!$C:$C),\"yyyy/MM/dd\")&\" - \"&TEXT(MAX(年度兌回!$D:$D),\"yyyy/MM/dd\")";
                    sheet2.Cell(6, 2).Style.DateFormat.Format = "yyyy/MM/dd"; // 欄位格式
                    sheet2.Cell(6, 2).FormulaA1 = "=MIN(年度兌回!$E:$E)";
                    sheet2.Cell(6, 3).Style.DateFormat.Format = "yyyy/MM/dd"; // 欄位格式
                    sheet2.Cell(6, 3).FormulaA1 = "=MAX(年度兌回!$F:$F)";
                    sheet2.Cell(6, 4).FormulaA1 = "=SUM(年度兌回!$G:$G)";
                    sheet2.Cell(6, 5).FormulaA1 = "=SUM(年度兌回!$H:$H)";
                    sheet2.Cell(6, 6).FormulaA1 = "=$E6/$D6";
                    sheet2.ColumnWidth += 8;
                    sheet2.Column("A").Width += 8;
                    sheet2.Range("A5:F6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    workbook.SaveAs(@"C:\Website\UpdateCouponYearly\file\" + _mall + "年度兌回.xlsx");  //正式
                    //workbook.SaveAs(@"E:\POJHIH\Website\UpdateCouponYearly\file\" + _mall + "年度兌回.xlsx");  //本地
                }
            }
            catch (Exception ex)
            {
                string Failedmessage = "年度遞延報表發生錯誤，錯誤訊息：" + ex;
                commonUtility.Txt(Failedmessage);
            }
        }
        /// <summary>
        /// 每月年度遞延報表
        /// </summary>
        public static void Deferred_Submit()
        {
            CommonUtility commonUtility = new CommonUtility();
            try
            {
                DateTime _sDate = new DateTime(DateTime.Now.Year, 1, 1);//年初
                DateTime _eDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);//當月初
                if (ConfigurationManager.AppSettings["DeBUG"] == "Y")
                {
                    DateTime.TryParseExact(ConfigurationManager.AppSettings["startDate"], "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateS);
                    _sDate = dateS;
                    DateTime.TryParseExact(ConfigurationManager.AppSettings["endDate"], "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateE);
                    _eDate = dateE;
                }
                DataTable _dt = CouponDeferred_Get(_sDate, _eDate.AddDays(1));
                List<string> _listMall = new List<string>() { "32", "34", "37", "40", "42", "48", "50", "51", "52", "53", "54", "55", "72" };
                //List<string> _listMall = new List<string>() { "53" };
                foreach (var _mall in _listMall)
                {
                    long _TotalPrice = 0;
                    long _UsedPrice = 0;
                    XLWorkbook workbook = new XLWorkbook();
                    var sheet = workbook.Worksheets.Add("兌出-兌回報表");
                    int _num = 6; int nPrice = 0; int uPrice = 0; int totalPrice = 0;
                    DataRow[] _drTemp = _dt.Select("MallId=" + _mall);
                    sheet.Cell(1, 1).Value = "查詢日期：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    sheet.Cell(2, 1).Value = "查詢人員：系統排程匯出";
                    sheet.Cell(3, 1).Value = "查詢公司：" + _mall;
                    sheet.Cell(4, 1).Value = $"兌回起訖：{_sDate.ToString("yyyy-MM-dd")}至{_eDate.ToString("yyyy-MM-dd")}";
                    sheet.Cell(5, 1).Value = "券年度";
                    sheet.Cell(5, 2).Value = "券代號";
                    sheet.Cell(5, 3).Value = "兌出起始";
                    sheet.Cell(5, 4).Value = "兌出結束";
                    sheet.Cell(5, 5).Value = "兌回起始";
                    sheet.Cell(5, 6).Value = "兌回結束";
                    sheet.Cell(5, 7).Value = "兌出(除稅)";
                    sheet.Cell(5, 8).Value = "兌回(除稅)";
                    sheet.Cell(5, 9).Value = "兌出-兌回(除稅)";
                    //sheet.Cell(5, 10).Value = "兌出業績(除稅)";
                    //sheet.Cell(5, 11).Value = "兌回業績(除稅)";
                    if (_drTemp.Count() > 0)
                    {

                        foreach (var item in _drTemp)
                        {
                            //long _total = _cs.CouponGive(Convert.ToInt32(item["GId"]), item["MallId"].ToString());
                            //_TotalPrice += _total;
                            sheet.Cell(_num, 1).Value = item["Year"].ToString();
                            sheet.Cell(_num, 2).Value = item["GiftsNo"].ToString();
                            sheet.Cell(_num, 3).Value = string.Format("{0:yyyy/MM/dd}", item["ExchangeStart"]);
                            sheet.Cell(_num, 4).Value = string.Format("{0:yyyy/MM/dd}", item["ExchangeEnd"]);
                            sheet.Cell(_num, 5).Value = string.Format("{0:yyyy/MM/dd}", item["UsedStart"]);
                            sheet.Cell(_num, 6).Value = string.Format("{0:yyyy/MM/dd}", item["UsedEnd"]);
                            sheet.Cell(_num, 7).Value = Convert.ToInt32(item["NPrice"]);
                            sheet.Cell(_num, 8).Value = Convert.ToInt32(item["UPrice"]);
                            sheet.Cell(_num, 9).Value = Convert.ToInt32(item["TotalPrice"]);
                            //sheet.Cell(_num, 10).Value = _total;
                            //sheet.Cell(_num, 11).Value = Convert.ToInt64(item["UsedPrice"]);
                            sheet.Range($"A{_num}:I{_num}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right; //靠右
                            _num++;
                            nPrice += Convert.ToInt32(item["NPrice"]);
                            uPrice += Convert.ToInt32(item["UPrice"]);
                            totalPrice += Convert.ToInt32(item["TotalPrice"]);
                            //_UsedPrice += Convert.ToInt64(item["UsedPrice"]);
                        }
                        sheet.Cell(_num, 6).Value = "總計";
                        sheet.Cell(_num, 7).FormulaA1 = $"=SUM(G6:G{_num - 1})";
                        sheet.Cell(_num, 8).FormulaA1 = $"=SUM(H6:H{_num - 1})";
                        sheet.Cell(_num, 9).FormulaA1 = $"=SUM(I6:I{_num - 1})";
                        //sheet.Cell(_num, 10).Value = _TotalPrice;
                        //sheet.Cell(_num, 11).Value = _UsedPrice;
                        sheet.ColumnWidth += 8;
                        

                    }

                    workbook.SaveAs(@"C:\Website\UpdateCouponYearly\file\" + _mall + "兌出-兌回.xlsx");  //正式
                    //workbook.SaveAs(@"E:\POJHIH\Website\UpdateCouponYearly\file\" + _mall + "兌出-兌回.xlsx");  //本地

                }
            }
            catch (Exception ex)
            {
                string Failedmessage = "遞延報表發生錯誤，錯誤訊息：" + ex;
                commonUtility.Txt(Failedmessage);
            }
        }
    }
}
