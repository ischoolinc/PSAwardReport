using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PSAwardReport
{
    
    public class AwardScore
    {

        /// <summary>
        /// 成績種類(期中總成績、期未總成績、進步分數、學期總成績)
        /// </summary>
        public string ScoreType { get; set; }

        /// <summary>
        /// 分數
        /// </summary>
        public decimal? Score { get; set; }

        /// <summary>
        /// 排名
        /// </summary>
        public string Rank { get; set; }

        


    }
}
