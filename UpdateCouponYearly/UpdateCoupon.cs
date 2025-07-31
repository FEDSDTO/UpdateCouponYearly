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
        public static DataTable CouponDeferred_Get()
        {
            DB_Connection _db = new DB_Connection();
            List<SqlParameter> _Parameter = new List<SqlParameter>();
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
            string _SQL = @"Select A.Year,A.MallId,A.GId,M.PosMallId+GiftsNo GiftsNo,E.ExchangeStart,E.ExchangeEnd,E.UsedStart,E.UsedEnd,
                                 Round(Sum(NPrice)/1.05,0) NPrice,Round(Sum(UPrice)/1.05,0) UPrice,Round(Sum(NPrice)/1.05,0) - Round(Sum(UPrice)/1.05,0) TotalPrice    
                             From (
            
                                       Select G.Year,C.MallId,GId,G.GiftsNo,Count(GId)*G.Denomination NPrice,0 UPrice,UsedStart,UsedEnd
                                       From Coupon C
                                        Left join(
                                                Select * 
                                                From Gifts
												Where IsUse = 1
                                                )G on C.GId = G.Id
                                     Where UsedEnd >= @SDate and UsedEnd < @EDate And G.Type = 'C' And Status != 'F' And C.CreateOn < @EDate And SAPType <> 'B' --And SUBSTRING(C.MemberId,1,1) <> 'B' And MemberId <> ''
                                       Group by C.MallId,GId,GiftsNo,UsedStart,UsedEnd,Denomination,G.Year
                                      union all

                                      Select G.Year,C.MallId,C.GId,G.GiftsNo,0 NPrice,Count(C.GId)*G.Denomination UPrice,C.UsedStart,C.UsedEnd
                                       From Coupon C
                                        Left join(
                                            Select * 
                                            From Gifts
											Where IsUse = 1
                                            )G on C.GId = G.Id
                                        Left join(
                                            Select *
                                            From UsedRule
											Where IsUse = 1
                                               ) UR on C.GId = UR.GId And C.MallId = UR.MallId
                                         Where C.UsedEnd >= @SDate and C.UsedEnd < @EDate And Status = 'U' And G.Type = 'C'  And C.UsedDate < @EDate And SAPType <> 'B'  And Len(C.MemberId) < 10 And MemberId <> ''
                                           Group by C.MallId,C.GId,GiftsNo,C.UsedStart,C.UsedEnd,Denomination,G.Year
                                      )A
                             Left join(
                                       Select * 
                                       From Mall
                                       ) M on A.MallId = M.MallId
                              Left join(
                                       Select * 
                                       From UsedRule
									   Where IsUse = 1
                                       ) E on A.GId = E.GId And A.MallId = E.MallId
									   --WHERE E.ExchangeStart IS NULL OR E.ExchangeEnd IS NULL OR ExchangeStart < '2019-01-01'
                             Group by A.MallId,A.GId,GiftsNo,PosMallId,A.UsedStart,A.UsedEnd,E.ExchangeStart,E.ExchangeEnd,E.UsedStart,E.UsedEnd,A.Year
							 order by MallId,UsedEnd
                                                ";
            _Parameter.Add(new SqlParameter("SDate", _sDate));
            _Parameter.Add(new SqlParameter("EDate", _eDate));

            return _db.GetDataTable(_SQL, _Parameter);
        }
        /// <summary>
        /// 年度遞延報表
        /// </summary>

        public static void Deferred_SubmitYearly()
        {
            CommonUtility commonUtility = new CommonUtility();
            try
            {
                //string _now = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddDays(-1).ToString("yyyyMMdd");
                DataTable _dt = CouponDeferred_Get();
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
                    sheet.Cell(1, 1).Value = "查詢日期：" + DateTime.Now.ToString("yyyy/MM/dd");
                    sheet.Cell(2, 1).Value = "查詢人員：系統排程匯出";
                    sheet.Cell(3, 1).FormulaA1 = "=\"兌回迄日：\"&TEXT(MAX($E:$E),\"yyyy/MM/dd\")";
                    sheet.Cell(4, 1).Value = "分公司：" + _mall;
                    sheet.Cell(5, 1).Value = "券代號";
                    sheet.Cell(5, 2).Value = "兌出起始";
                    sheet.Cell(5, 3).Value = "兌出結束";
                    sheet.Cell(5, 4).Value = "兌回起始";
                    sheet.Cell(5, 5).Value = "兌回結束";
                    sheet.Cell(5, 6).Value = "兌出(除稅)";
                    sheet.Cell(5, 7).Value = "兌回(除稅)";
                    sheet.Cell(5, 8).Value = "兌出-兌回(除稅)";
                    if (_drTemp.Count() > 0)
                    {

                        foreach (var item in _drTemp)
                        {
                            sheet.Cell(_num, 1).Value = item["GiftsNo"].ToString();
                            DateTime.TryParse(string.Format("{0:yyyy/MM/dd}", item["ExchangeStart"]), out DateTime ExchangeStart); //string 轉成datetime
                            sheet.Cell(_num, 2).Value = ExchangeStart;
                            DateTime.TryParse(string.Format("{0:yyyy/MM/dd}", item["ExchangeEnd"]), out DateTime ExchangeEnd); //string 轉成datetime
                            sheet.Cell(_num, 3).Value = ExchangeEnd;
                            DateTime.TryParse(string.Format("{0:yyyy/MM/dd}", item["UsedStart"]), out DateTime UsedStart); //string 轉成datetime
                            sheet.Cell(_num, 4).Value = UsedStart;
                            DateTime.TryParse(string.Format("{0:yyyy/MM/dd}", item["UsedEnd"]), out DateTime UsedEnd); //string 轉成datetime
                            sheet.Cell(_num, 5).Value = UsedEnd;
                            sheet.Cell(_num, 6).Value = Convert.ToInt32(item["NPrice"]);
                            sheet.Cell(_num, 7).Value = Convert.ToInt32(item["UPrice"]);
                            sheet.Cell(_num, 8).Value = Convert.ToInt32(item["TotalPrice"]);

                            _num++;
                        }
                        sheet.ColumnWidth += 8;

                    }

                    var sheet2 = workbook.Worksheets.Add("兌回率計算");
                    sheet2.Cell(1, 1).Value = "查詢日期：" + DateTime.Now.ToString("yyyy/MM/dd");
                    sheet2.Cell(2, 1).Value = "查詢人員：系統排程匯出";
                    sheet2.Cell(3, 1).FormulaA1 = "=年度兌回!A3";
                    sheet2.Cell(4, 1).Value = "分公司：" + _mall;
                    sheet2.Cell(5, 1).Value = "兌出期間";
                    sheet2.Cell(5, 2).Value = "兌回起日";
                    sheet2.Cell(5, 3).Value = "兌回迄日";
                    sheet2.Cell(5, 4).Value = "兌出";
                    sheet2.Cell(5, 5).Value = "兌回";
                    sheet2.Cell(5, 6).Value = "兌回率";
                    sheet2.Cell(6, 1).FormulaA1 = "=TEXT(MIN(年度兌回!$B:$B),\"yyyy/MM/dd\")&\" - \"&TEXT(MAX(年度兌回!$C:$C),\"yyyy/MM/dd\")";
                    sheet2.Cell(6, 2).Style.DateFormat.Format = "yyyy/MM/dd"; // 欄位格式
                    sheet2.Cell(6, 2).FormulaA1 = "=MIN(年度兌回!$D:$D)";
                    sheet2.Cell(6, 3).Style.DateFormat.Format = "yyyy/MM/dd"; // 欄位格式
                    sheet2.Cell(6, 3).FormulaA1 = "=MAX(年度兌回!$E:$E)";
                    sheet2.Cell(6, 4).FormulaA1 = "=SUM(年度兌回!$F:$F)";
                    sheet2.Cell(6, 5).FormulaA1 = "=SUM(年度兌回!$G:$G)";
                    sheet2.Cell(6, 6).FormulaA1 = "=$E6/$D6";
                    sheet2.ColumnWidth += 8;
                    sheet2.Column("A").Width += 8;
                    sheet2.Range("A5:F6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    //workbook.SaveAs(@"C:\Website\UpdateCouponYearly\file\" + _mall + "年度兌回.xlsx");  //正式
                    workbook.SaveAs(@"E:\POJHIH\Website\UpdateCouponYearly\" + _mall + "年度兌回.xlsx");  //本地
                }
            }
            catch (Exception ex)
            {
                string Failedmessage = "年度遞延報表發生錯誤，錯誤訊息：" + ex;
                commonUtility.Txt(Failedmessage);
            }
        }
    }
}
