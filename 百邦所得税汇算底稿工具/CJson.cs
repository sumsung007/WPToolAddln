using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace 百邦所得税汇算底稿工具
{
    public class CJson
    {
        public string zfbz {get;set;}
        public string queryType {get;set; }
        public voHj2[] voHj2 { get; set; }
        public string msg { get; set; }
        public voHj1[] voHj1 { get; set; }
        public lsWssbjl[] lsWssbjl { get; set; }
        public string sbyear { get; set; }
        public string sbrq_month { get; set; }
        public string cxtj { get; set; }
        public string sbrq_year { get; set; }
        public string fycx { get; set; }
        public decimal sjyjjeHj { get; set; }
        public pagination pagination { get; set; }
        public Boolean success { get; set; }
    }
    public class voHj2
    { }
    public class voHj1
    { }

    public class pagination
    {
        public int pageSize { get; set; }
        public int pageNumber { get; set; }
        public int totalCount { get; set; }
        public int pageCount { get; set; }
    }

    public class lsWssbjl
    {
        public string skssqq { get; set; }//税款所属期起
        public string zspm_dm { get; set; }//征收品目
        public string zszm_dm { get; set; }//征收子目
        public string ynse { get; set; }//应纳税额
        public string jsyj { get; set; }//计税金额
        public string zsxm_dm { get; set; }//征收项目
        public string nssbrq { get; set; }//纳税申报日期
        public string sl_1 { get; set; }//税率
        public string rkrq { get; set; }//入库日期
        public string ybtse { get; set; }//实际应缴税额
        public string skssqz { get; set; }//税款所属期止
        public string jmse { get; set; }//减免税额
        public string yjse { get; set; }//已缴税额

    }

}
