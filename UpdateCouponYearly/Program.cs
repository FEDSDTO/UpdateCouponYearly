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
            //每年1月1日
            if (DateTime.Now.Month == 1 && DateTime.Now.Day == 1)
            {
                UpdateCoupon.Deferred_SubmitYearly();
                CommonUtility.MoveFiles();
                if (ConfigurationManager.AppSettings["SendMail"] == "Y")
                {
                    CommonUtility.SendMail();
                } 
            }
            //人工處裡
            if (ConfigurationManager.AppSettings["DeBUG"] == "Y")
            {
                UpdateCoupon.Deferred_SubmitYearly();
                CommonUtility.MoveFiles();
            }
        }
    }
}
