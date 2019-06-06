using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA;
using FISCA.Presentation;
using K12.Presentation;
using FISCA.Permission;
using JHSchool;


namespace PSAwardReport
{
    public class Program
    {

        [FISCA.MainMethod()]
        public static void Main()
        {
            FISCA.UDT.AccessHelper accessHelper = new FISCA.UDT.AccessHelper();

            {
                Catalog ribbon = RoleAclSource.Instance["班級"]["報表"];
                ribbon.Add(new RibbonFeature("592F43FD-73C0-47F7-A259-A2A0A2F0F99E", "五育獎報表"));

                MotherForm.RibbonBarItems["班級", "資料統計"]["報表"]["成績相關報表"]["五育獎報表"].Enable = UserAcl.Current["592F43FD-73C0-47F7-A259-A2A0A2F0F99E"].Executable;

                MotherForm.RibbonBarItems["班級", "資料統計"]["報表"]["成績相關報表"]["五育獎報表"].Click += delegate
                {


                };
            }


            {
                Catalog ribbon = RoleAclSource.Instance["班級"]["報表"];
                ribbon.Add(new RibbonFeature("046FD063-576D-483C-8749-325D43748D11", "進步獎報表"));

                MotherForm.RibbonBarItems["班級", "資料統計"]["報表"]["成績相關報表"]["進步獎報表"].Enable = UserAcl.Current["046FD063-576D-483C-8749-325D43748D11"].Executable;

                MotherForm.RibbonBarItems["班級", "資料統計"]["報表"]["成績相關報表"]["進步獎報表"].Click += delegate
                {


                };
            }



        }
    }
}
