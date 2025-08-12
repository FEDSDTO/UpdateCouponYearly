using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UpdateCouponYearly
{
    class Program
    {
        static void Main(string[] args)
        {
            //每月1日
            if (DateTime.Now.Day == 1)
            {
                UpdateCoupon.Deferred_Submit();
                CommonUtility.MoveFiles();
                if (ConfigurationManager.AppSettings["SendMail"] == "Y")
                {
                    CommonUtility.SendMail();
                }
            }
            //人工處裡
            if (ConfigurationManager.AppSettings["DeBUG"] == "Y")
            {
                if (ConfigurationManager.AppSettings["DeBugYear"] == "Y")
                {
                    UpdateCoupon.Deferred_SubmitYearly();
                }
                else
                {
                    UpdateCoupon.Deferred_Submit();
                }
                CommonUtility.MoveFiles();
            }
        }
    }
}
