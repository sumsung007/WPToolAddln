using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace 百邦所得税汇算底稿工具
{
    class Codes2018
    {

        public static void 报表填写()
        {
            if (CU.文件判断())
            {

                if ((WorkingPaper.Wb.Sheets["余额表"].Range["A5"] != null) &&
                    (Globals.WPToolAddln.Application.ActiveSheet.Name == "余额表") &&
                    (MessageBox.Show("是否填写报表？", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes))
                {
                    if (true)
                    {
                        //try
                        {
                            string kemu, daima;
                            WorkingPaper.Wb.Application.ScreenUpdating = false;
                            WorkingPaper.Wb.Sheets["资产负债表"].Range["C6:D15,C18:D22,C24:D31,G5:H15,G18:H21,G28:H31"]
                                .Value = "";
                            WorkingPaper.Wb.Sheets["利润表"].Range["C6:F7,C9:F17,C19:F20,C22:F22,E24:F24,C25:F26,C28:F33,C35:F40"]
                                .Value = "";
                            Worksheet SH = WorkingPaper.Wb.Sheets["余额表"];
                            int N = SH.Cells[SH.UsedRange.Rows.Count + 1, 2].End[XlDirection.xlUp].Row;
                            int changdu = (int)WorkingPaper.Wb.Sheets["辅助表"].Range["B16"].Value;// 一级科目长度
                            object[,] YEB = SH.Range["A5:H" + N.ToString()].Value2;
                            double[] qc = new double[60], qm = new double[60], lrb = new double[18];
                            N = N - 4;
                            for (int i = 1; i <= N - 1; i++)
                            {
                                if ((CU.Zifu(YEB[i, 1]).ToString().Trim().Length == changdu)
                                    && YEB[i, 2] != null)
                                {
                                    kemu = CU.Zifu(YEB[i, 2]);
                                    daima = CU.Zifu(YEB[i, 1]);
                                    switch (kemu)
                                    {
                                        case "现金":
                                        case "库存现金":
                                            qc[0] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[0] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "银行存款":
                                            qc[1] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[1] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "其他货币资金":
                                            qc[2] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[2] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "交易性金融资产":
                                            qc[3] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[3] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "短期投资":
                                            qc[4] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[4] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "应收票据":
                                            qc[5] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[5] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "应收股息":
                                        case "应收股利":
                                            qc[8] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[8] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "应收利息":
                                            qc[9] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[9] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "应收帐款":
                                        case "应收账款":
                                            qc[6] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[6] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "预付帐款":
                                        case "预付账款":
                                            qc[7] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[7] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "其他应收款":
                                            qc[10] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[10] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;

                                        case "存货":
                                            qc[11] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[11] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "库存商品":
                                            qc[12] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[12] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "原材料":
                                            qc[13] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[13] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;

                                        case "产成品":
                                            qc[14] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[14] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "生产成本":
                                            qc[15] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[15] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "制造费用":
                                            qc[16] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[16] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "工程施工":
                                            qc[17] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[17] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "委托加工物质":
                                            qc[18] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[18] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "材料成本差异":
                                            qc[19] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[19] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "材料采购":
                                            qc[20] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[20] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "在途物质":
                                            qc[21] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[21] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "发出商品":
                                            qc[22] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[22] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;

                                        case "应收出口退税":
                                            qc[23] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[23] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;

                                        case "待摊费用":
                                            qc[24] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[24] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "长期股权投资":
                                            qc[26] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[26] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "持有至到期投资":
                                            qc[27] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[27] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "可供出售金融资产":
                                            qc[28] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[28] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "长期债权投资":
                                            qc[25] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[25] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "固定资产":
                                            qc[29] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[29] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "累计折旧":
                                            qc[30] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[30] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "工程物资":
                                            qc[31] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[31] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "在建工程":
                                            qc[32] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[32] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "固定资产清理":
                                            qc[33] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[33] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "生产性生物资产":
                                            qc[34] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[34] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "无形资产":
                                            qc[35] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[35] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "开发支出":
                                        case "研发支出":
                                            qc[36] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[36] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "长期待摊费用":
                                            qc[37] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[37] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;

                                        case "短期借款":
                                            qc[38] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[38] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "应付票据":
                                            qc[39] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[39] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "应付帐款":
                                        case "应付账款":
                                            qc[40] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[40] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "预收帐款":
                                        case "预收账款":
                                            qc[41] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[41] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "应付工资":
                                        case "应付职工薪酬":
                                            qc[42] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[42] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "应付福利费":
                                            qc[43] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[43] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "应付利润":
                                            qc[44] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[44] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "应交税金":
                                        case "应交税费":
                                            qc[45] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[45] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "其他应交款":
                                            qc[46] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[46] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "应付利息":
                                            qc[47] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[47] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "其他应付款":
                                            qc[48] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[48] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "预提费用":
                                            qc[49] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[49] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "长期借款":
                                            qc[50] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[50] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "长期应付款":
                                            qc[51] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[51] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;

                                        case "资本公积":
                                            qc[52] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[52] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "盈余公积":
                                            qc[53] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[53] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "本年利润":
                                            qc[54] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[54] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            break;
                                        case "利润分配":
                                            qc[55] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                            qm[55] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            int X = i + 1;
                                            while ((X <= N - 2) && (YEB[X, 1].ToString().Contains(daima)))
                                            {
                                                if (YEB[X, 2].ToString().Contains("未分配利润"))
                                                {
                                                    qc[55] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                                    qm[55] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                                    break;
                                                }
                                                X++;
                                            }
                                            break;
                                        //资产负债表到此结束

                                        //利润表开始
                                        case "主营业务收入":
                                        case "商品销售收入":
                                        case "产品销售收入":
                                            lrb[0] = Math.Round(CU.Shuzi(YEB[i, 6]), 2);
                                            break;
                                        case "主营业务成本":
                                        case "商品销售成本":
                                        case "产品销售成本":
                                            lrb[1] = Math.Round(CU.Shuzi(YEB[i, 5]), 2);
                                            break;
                                        case "主营业务税金及附加":
                                        case "商品销售税金及附加":
                                        case "产品销售税金及附加":
                                        case "营业税金及附加":
                                            lrb[2] = Math.Round(CU.Shuzi(YEB[i, 5]), 2);
                                            break;
                                        case "其他业务收入":
                                            lrb[3] = Math.Round(CU.Shuzi(YEB[i, 6]), 2);
                                            break;
                                        case "其他业务支出":
                                            lrb[4] = Math.Round(CU.Shuzi(YEB[i, 5]), 2);
                                            break;
                                        case "营业费用":
                                        case "销售费用":
                                        case "经营费用":
                                            lrb[5] = Math.Round(CU.Shuzi(YEB[i, 5]), 2);
                                            break;
                                        case "管理费用":
                                            lrb[6] = Math.Round(CU.Shuzi(YEB[i, 5]), 2);
                                            break;
                                        case "财务费用":
                                            lrb[7] = Math.Round(CU.Shuzi(YEB[i, 5]), 2);
                                            break;
                                        case "投资收益":
                                        case "投资利润":
                                            lrb[8] = Math.Round(CU.Shuzi(YEB[i, 6]), 2);
                                            break;
                                        case "营业外收入":
                                        case "补贴收入":
                                            lrb[9] = Math.Round(CU.Shuzi(YEB[i, 6]), 2);
                                            break;
                                        case "营业外支出":
                                            lrb[10] = Math.Round(CU.Shuzi(YEB[i, 5]), 2);
                                            break;
                                        case "所得税":
                                        case "所得税费用":
                                            lrb[11] = Math.Round(CU.Shuzi(YEB[i, 5]), 2);
                                            break;

                                        case "资产减值损失":
                                            lrb[12] = Math.Round(CU.Shuzi(YEB[i, 5]), 2);
                                            break;
                                        case "公允价值变动收益":
                                        case "公允价值变动损益":
                                            lrb[13] = Math.Round(CU.Shuzi(YEB[i, 6]), 2);
                                            break;

                                        default:
                                            if (Regex.IsMatch(kemu, ".*低值易耗品"))
                                            {
                                                qc[57] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                                qm[57] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                            }
                                            else
                                            {
                                                if (Regex.IsMatch(kemu, ".*包装物"))
                                                {
                                                    qc[58] = Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                                    qm[58] = Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                                }
                                                else
                                                {
                                                    if (Regex.IsMatch(kemu, ".*(资本|股本)"))
                                                    {
                                                        qc[59] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]),
                                                            2);
                                                        qm[59] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]),
                                                            2);
                                                    }
                                                }
                                            }
                                            break;
                                    }
                                }
                            }
                            double[,] ldzc = new double[10, 2];
                            ldzc[0, 1] = qc[0] + qc[1] + qc[2]; //货币资金
                            ldzc[0, 0] = qm[0] + qm[1] + qm[2];
                            ldzc[1, 1] = qc[3] + qc[4]; //短期投资
                            ldzc[1, 0] = qm[3] + qm[4];
                            ldzc[2, 1] = qc[5]; //应收票据
                            ldzc[2, 0] = qm[5];
                            ldzc[3, 1] = qc[6]; //应收账款
                            ldzc[3, 0] = qm[6];
                            ldzc[4, 1] = qc[7]; //预付账款
                            ldzc[4, 0] = qm[7];
                            ldzc[5, 1] = qc[8]; //应收股息
                            ldzc[5, 0] = qm[8];
                            ldzc[6, 1] = qc[9]; //应收利息
                            ldzc[6, 0] = qm[9];
                            ldzc[7, 1] = qc[10]; //其他应收款
                            ldzc[7, 0] = qm[10];
                            ldzc[8, 1] = qc[11] + qc[12] + qc[13] + qc[14] + qc[15] + qc[16] + qc[17] + qc[18] +
                                         qc[19] + qc[20] + qc[21] + qc[22] + qc[57] + qc[58]; //存货
                            ldzc[8, 0] = qm[11] + qm[12] + qm[13] + qm[14] + qm[15] + qm[16] + qm[17] + qm[18] +
                                         qm[19] + qm[20] + qm[21] + qm[22] + qm[57] + qm[58];
                            ldzc[9, 1] = qc[23] + qc[24]; //其他流动资产=待摊费用+应收出口退税
                            ldzc[9, 0] = qm[23] + qm[24];


                            WorkingPaper.Wb.Sheets["资产负债表"].Range["C6:D15"].Value2 = ldzc;

                            ldzc = new double[5, 2];
                            ldzc[0, 1] = qc[25]; //长期债权投资
                            ldzc[0, 0] = qm[25];
                            ldzc[1, 1] = qc[26] + qc[27] + qc[28]; //长期股权投资
                            ldzc[1, 0] = qm[26] + qm[27] + qm[28];
                            ldzc[2, 1] = qc[29]; //固定资产
                            ldzc[2, 0] = qm[29];
                            ldzc[3, 1] = qc[30]; //累计折旧
                            ldzc[3, 0] = qm[30];
                            ldzc[4, 1] = 0; //减值准备
                            ldzc[4, 0] = 0;
                            WorkingPaper.Wb.Sheets["资产负债表"].Range["C18:D22"].Value2 = ldzc;

                            ldzc = new double[8, 2];
                            ldzc[0, 1] = qc[32]; //在建工程
                            ldzc[0, 0] = qm[32];
                            ldzc[1, 1] = qc[31]; //工程物资
                            ldzc[1, 0] = qm[31];
                            ldzc[2, 1] = qc[33]; //固定资产清理
                            ldzc[2, 0] = qm[33];
                            ldzc[3, 1] = qc[34]; //生产性生物资产
                            ldzc[3, 0] = qm[34];
                            ldzc[4, 1] = qc[35]; //无形资产
                            ldzc[4, 0] = qm[35];
                            ldzc[5, 1] = qc[36]; //开发支出
                            ldzc[5, 0] = qm[36];
                            ldzc[6, 1] = qc[37]; //长期待摊费用
                            ldzc[6, 0] = qm[37];
                            ldzc[7, 1] = 0; //其他非流动资产
                            ldzc[7, 0] = 0;
                            WorkingPaper.Wb.Sheets["资产负债表"].Range["C24:D31"].Value2 = ldzc;

                            ldzc = new double[10, 2];
                            ldzc[0, 1] = qc[38]; //短期借款
                            ldzc[0, 0] = qm[38];
                            ldzc[1, 1] = qc[39]; //应付票据
                            ldzc[1, 0] = qm[39];
                            ldzc[2, 1] = qc[40]; //应付账款
                            ldzc[2, 0] = qm[40];
                            ldzc[3, 1] = qc[41]; //预收账款
                            ldzc[3, 0] = qm[41];
                            ldzc[4, 1] = qc[42] + qc[43]; //应付职工薪酬=应付工资+应付福利费
                            ldzc[4, 0] = qm[42] + qm[43];
                            ldzc[5, 1] = qc[45] + qc[46]; //应交税费=应交税金+其他应交款
                            ldzc[5, 0] = qm[45] + qm[46];
                            ldzc[6, 1] = qc[47]; //应付利息
                            ldzc[6, 0] = qm[47];
                            ldzc[7, 1] = qc[44]; //应付利润
                            ldzc[7, 0] = qm[44];
                            ldzc[8, 1] = qc[48]; //其他应付款
                            ldzc[8, 0] = qm[48];
                            ldzc[9, 1] = qc[49]; //其他流动负债 含预提费用
                            ldzc[9, 0] = qm[49];
                            WorkingPaper.Wb.Sheets["资产负债表"].Range["G6:H15"].Value2 = ldzc;


                            ldzc = new double[4, 2];
                            ldzc[0, 1] = qc[50]; //长期借款
                            ldzc[0, 0] = qm[50];
                            ldzc[1, 1] = qc[51]; //长期应付款
                            ldzc[1, 0] = qm[51];

                            WorkingPaper.Wb.Sheets["资产负债表"].Range["G18:H21"].Value2 = ldzc;

                            ldzc = new double[4, 2];
                            ldzc[0, 1] = qc[59]; //实收资本
                            ldzc[0, 0] = qm[59];
                            ldzc[1, 1] = qc[52]; //资本公积
                            ldzc[1, 0] = qm[52];
                            ldzc[2, 1] = qc[53]; //盈余公积
                            ldzc[2, 0] = qm[53];
                            ldzc[3, 1] = qc[54] + qc[55]; //未分配利润+本年利润
                            ldzc[3, 0] = qm[54] + qm[55];

                            WorkingPaper.Wb.Sheets["资产负债表"].Range["G28:H31"].Value2 = ldzc;

                            ldzc = new double[2, 1];
                            ldzc[0, 0] = lrb[0]; //主营业务收入
                            ldzc[1, 0] = lrb[3]; //其他业务收入
                            WorkingPaper.Wb.Sheets["利润表"].Range["C6:C7"].Value2 = ldzc;

                            ldzc = new double[9, 1];
                            ldzc[0, 0] = lrb[1]; //主营业务成本
                            ldzc[1, 0] = lrb[4]; //其他业务成本
                            ldzc[2, 0] = lrb[2]; //营业税金及附加
                            ldzc[3, 0] = lrb[5]; //销售费用
                            ldzc[4, 0] = lrb[6]; //管理费用
                            ldzc[5, 0] = lrb[7]; //财务费用
                            ldzc[6, 0] = lrb[12]; //资产减值损失
                            ldzc[7, 0] = lrb[13]; //公允价值变动损益
                            ldzc[8, 0] = lrb[8]; //投资收益
                            WorkingPaper.Wb.Sheets["利润表"].Range["C9:C17"].Value2 = ldzc;

                            ldzc = new double[2, 1];
                            ldzc[0, 0] = lrb[9]; //营业外收入
                            ldzc[1, 0] = lrb[10]; //营业外支出
                            WorkingPaper.Wb.Sheets["利润表"].Range["C19:C20"].Value2 = ldzc;

                            WorkingPaper.Wb.Sheets["利润表"].Cells[22, 3].Value = lrb[11]; //所得税
                            WorkingPaper.Wb.Application.ScreenUpdating = true;
                        }
                        //catch (Exception ex)
                        //{
                        //    WorkingPaper.Wb.Application.ScreenUpdating = true;
                        //    MessageBox.Show("用户操作出现错误：" + ex.Message);
                        //}
                        if (CU.Shuzi(WorkingPaper.Wb.Sheets["资产负债表"].Range["C35"].Value2) != 0)
                        {
                            MessageBox.Show("报表填写完毕，请复查!" + "资产负债表未分配利润期末余额与利润表期末未分配利润差异" +
                                            CU.Shuzi(WorkingPaper.Wb.Sheets["资产负债表"].Range["C35"].Value2).ToString("N"));
                        }
                        else
                            MessageBox.Show("报表填写完毕，请复查!");
                    }
                }
            }
        }

        public static void 底稿填写()
        {
            if (!CU.文件判断())
                return;
            string str;
            if (true)
            {
                //自动填表开始
                try
                {
                    WorkingPaper.Wb.Application.ScreenUpdating = false;           //自动填写底稿包含货币资金、往来、费用底稿
                    Worksheet SH = WorkingPaper.Wb.Sheets["余额表"];
                    int N = SH.Cells[SH.UsedRange.Rows.Count + 1, 2].End[XlDirection.xlUp].Row;
                    //SH.Columns[9].Clear();
                    //SH.Range["I2:I" + N].FormulaR1C1 = "= COUNTIF(C[-8], RC[-8] & \"?*\")";
                    object[,] YEB = SH.Range["A5:H" + N.ToString()].Value2;
                    object[,] Kemu = WorkingPaper.Wb.Sheets["辅助表"].Range["B17:B29"].Value2;
                    int n = 0;

                    if (CU.Shuzi(Kemu[1, 1]) != 0)                                     //现金
                    {
                        n = Convert.ToInt16(Kemu[1, 1]) - 1;                  //获取一级科目起始行
                        WorkingPaper.Wb.Worksheets["货币资金"].Cells[7, 4].Value = Math.Round(CU.Shuzi(YEB[n, 7]) - CU.Shuzi(YEB[n, 8]), 2);

                    }

                    if (CU.Shuzi(Kemu[2, 1]) != 0)                                        //银行存款
                    {
                        n = Convert.ToInt16(Kemu[2, 1]) - 1;
                        str = CU.Zifu(YEB[n, 1]);
                        int i = 13;
                        double m, sum = Math.Round(CU.Shuzi(YEB[n, 7]) - CU.Shuzi(YEB[n, 8]), 2);
                        do
                        {
                            if ((n == N) || (!CU.Zifu(YEB[n + 1, 1]).Contains(CU.Zifu(YEB[n, 1]))))
                            {
                                m = Math.Round(CU.Shuzi(YEB[n, 7]) - CU.Shuzi(YEB[n, 8]), 2);
                                WorkingPaper.Wb.Worksheets["货币资金"].Cells[i, 4].Value = m;
                                WorkingPaper.Wb.Worksheets["货币资金"].Cells[i, 2].Value = YEB[n, 2].ToString().Trim();
                                sum = sum - m;
                                i = i + 1;
                            }
                            n = n + 1;
                        } while (CU.Zifu(YEB[n, 1]).Contains(str) && (i <= 18));
                        if (i == 19)
                            WorkingPaper.Wb.Worksheets["货币资金"].Cells[i, 4].Value = sum;
                    }

                    if (CU.Shuzi(Kemu[3, 1]) != 0)                                //应收账款
                    {
                        n = Convert.ToInt16(Kemu[3, 1]) - 1;
                        str = CU.Zifu(YEB[n, 1]);
                        List<string> Mingcheng = new List<string>();
                        List<double> Jiefang = new List<double>();
                        List<double> Daifang = new List<double>();

                        do
                        {
                            if ((n == N) || (!CU.Zifu(YEB[n + 1, 1]).Contains(CU.Zifu(YEB[n, 1]))))
                            {
                                Mingcheng.Add(YEB[n, 2].ToString().Trim());
                                Jiefang.Add(CU.Shuzi(YEB[n, 7]));
                                Daifang.Add(CU.Shuzi(YEB[n, 8]));
                            }
                            n = n + 1;
                        } while (CU.Zifu(YEB[n, 1]).Contains(str));
                        string[,] mingcheng = new string[Mingcheng.Count, 1];
                        double[,] jiefang = new double[Jiefang.Count, 1];
                        double[,] daifang = new double[Daifang.Count, 1];

                        int k = 0;
                        foreach (string s in Mingcheng)
                        {
                            mingcheng[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Jiefang)
                        {
                            jiefang[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Daifang)
                        {
                            daifang[k, 0] = s;
                            k++;
                        }

                        WorkingPaper.Wb.Sheets["应收"].Range["A16:A" + (15 + Mingcheng.Count).ToString()].Value2 = mingcheng;
                        WorkingPaper.Wb.Sheets["应收"].Range["B16:B" + (15 + Jiefang.Count).ToString()].Value2 = jiefang;
                        WorkingPaper.Wb.Sheets["应收"].Range["C16:C" + (15 + Daifang.Count).ToString()].Value2 = daifang;
                    }

                    if (CU.Shuzi(Kemu[4, 1]) != 0)                                //预付账款
                    {
                        n = Convert.ToInt16(Kemu[4, 1]) - 1;
                        str = CU.Zifu(YEB[n, 1]);
                        List<string> Mingcheng = new List<string>();
                        List<double> Jiefang = new List<double>();
                        List<double> Daifang = new List<double>();

                        do
                        {
                            if ((n == N) || (!CU.Zifu(YEB[n + 1, 1]).Contains(CU.Zifu(YEB[n, 1]))))
                            {
                                Mingcheng.Add(YEB[n, 2].ToString().Trim());
                                Jiefang.Add(CU.Shuzi(YEB[n, 7]));
                                Daifang.Add(CU.Shuzi(YEB[n, 8]));
                            }
                            n = n + 1;
                        } while (CU.Zifu(YEB[n, 1]).Contains(str));
                        string[,] mingcheng = new string[Mingcheng.Count, 1];
                        double[,] jiefang = new double[Jiefang.Count, 1];
                        double[,] daifang = new double[Daifang.Count, 1];

                        int k = 0;
                        foreach (string s in Mingcheng)
                        {
                            mingcheng[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Jiefang)
                        {
                            jiefang[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Daifang)
                        {
                            daifang[k, 0] = s;
                            k++;
                        }

                        WorkingPaper.Wb.Sheets["预付"].Range["A16:A" + (15 + Mingcheng.Count).ToString()].Value2 = mingcheng;
                        WorkingPaper.Wb.Sheets["预付"].Range["B16:B" + (15 + Jiefang.Count).ToString()].Value2 = jiefang;
                        WorkingPaper.Wb.Sheets["预付"].Range["C16:C" + (15 + Daifang.Count).ToString()].Value2 = daifang;
                    }

                    if (CU.Shuzi(Kemu[5, 1]) != 0)                                //其他应收款
                    {
                        n = Convert.ToInt16(Kemu[5, 1]) - 1;
                        str = CU.Zifu(YEB[n, 1]);
                        List<string> Mingcheng = new List<string>();
                        List<double> Jiefang = new List<double>();
                        List<double> Daifang = new List<double>();

                        do
                        {
                            if ((n == N) || (!CU.Zifu(YEB[n + 1, 1]).Contains(CU.Zifu(YEB[n, 1]))))
                            {
                                Mingcheng.Add(YEB[n, 2].ToString().Trim());
                                Jiefang.Add(CU.Shuzi(YEB[n, 7]));
                                Daifang.Add(CU.Shuzi(YEB[n, 8]));
                            }
                            n = n + 1;
                        } while (CU.Zifu(YEB[n, 1]).Contains(str));
                        string[,] mingcheng = new string[Mingcheng.Count, 1];
                        double[,] jiefang = new double[Jiefang.Count, 1];
                        double[,] daifang = new double[Daifang.Count, 1];

                        int k = 0;
                        foreach (string s in Mingcheng)
                        {
                            mingcheng[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Jiefang)
                        {
                            jiefang[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Daifang)
                        {
                            daifang[k, 0] = s;
                            k++;
                        }

                        WorkingPaper.Wb.Sheets["其他应收"].Range["A16:A" + (15 + Mingcheng.Count).ToString()].Value2 = mingcheng;
                        WorkingPaper.Wb.Sheets["其他应收"].Range["B16:B" + (15 + Jiefang.Count).ToString()].Value2 = jiefang;
                        WorkingPaper.Wb.Sheets["其他应收"].Range["C16:C" + (15 + Daifang.Count).ToString()].Value2 = daifang;
                    }

                    if (CU.Shuzi(Kemu[6, 1]) != 0)                                //应付账款
                    {
                        n = Convert.ToInt16(Kemu[6, 1]) - 1;
                        str = CU.Zifu(YEB[n, 1]);
                        List<string> Mingcheng = new List<string>();
                        List<double> Jiefang = new List<double>();
                        List<double> Daifang = new List<double>();

                        do
                        {
                            if ((n == N) || (!CU.Zifu(YEB[n + 1, 1]).Contains(CU.Zifu(YEB[n, 1]))))
                            {
                                Mingcheng.Add(YEB[n, 2].ToString().Trim());
                                Jiefang.Add(CU.Shuzi(YEB[n, 7]));
                                Daifang.Add(CU.Shuzi(YEB[n, 8]));
                            }
                            n = n + 1;
                        } while (CU.Zifu(YEB[n, 1]).Contains(str));
                        string[,] mingcheng = new string[Mingcheng.Count, 1];
                        double[,] jiefang = new double[Jiefang.Count, 1];
                        double[,] daifang = new double[Daifang.Count, 1];

                        int k = 0;
                        foreach (string s in Mingcheng)
                        {
                            mingcheng[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Jiefang)
                        {
                            jiefang[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Daifang)
                        {
                            daifang[k, 0] = s;
                            k++;
                        }

                        WorkingPaper.Wb.Sheets["应付"].Range["A14:A" + (13 + Mingcheng.Count).ToString()].Value2 = mingcheng;
                        WorkingPaper.Wb.Sheets["应付"].Range["B14:B" + (13 + Jiefang.Count).ToString()].Value2 = jiefang;
                        WorkingPaper.Wb.Sheets["应付"].Range["C14:C" + (13 + Daifang.Count).ToString()].Value2 = daifang;
                    }

                    if (CU.Shuzi(Kemu[7, 1]) != 0)                                //预收账款
                    {
                        n = Convert.ToInt16(Kemu[7, 1]) - 1;
                        str = CU.Zifu(YEB[n, 1]);
                        List<string> Mingcheng = new List<string>();
                        List<double> Jiefang = new List<double>();
                        List<double> Daifang = new List<double>();

                        do
                        {
                            if ((n == N) || (!CU.Zifu(YEB[n + 1, 1]).Contains(CU.Zifu(YEB[n, 1]))))
                            {
                                Mingcheng.Add(YEB[n, 2].ToString().Trim());
                                Jiefang.Add(CU.Shuzi(YEB[n, 7]));
                                Daifang.Add(CU.Shuzi(YEB[n, 8]));
                            }
                            n = n + 1;
                        } while (CU.Zifu(YEB[n, 1]).Contains(str));
                        string[,] mingcheng = new string[Mingcheng.Count, 1];
                        double[,] jiefang = new double[Jiefang.Count, 1];
                        double[,] daifang = new double[Daifang.Count, 1];

                        int k = 0;
                        foreach (string s in Mingcheng)
                        {
                            mingcheng[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Jiefang)
                        {
                            jiefang[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Daifang)
                        {
                            daifang[k, 0] = s;
                            k++;
                        }

                        WorkingPaper.Wb.Sheets["预收"].Range["A14:A" + (13 + Mingcheng.Count).ToString()].Value2 = mingcheng;
                        WorkingPaper.Wb.Sheets["预收"].Range["B14:B" + (13 + Jiefang.Count).ToString()].Value2 = jiefang;
                        WorkingPaper.Wb.Sheets["预收"].Range["C14:C" + (13 + Daifang.Count).ToString()].Value2 = daifang;
                    }

                    if (CU.Shuzi(Kemu[8, 1]) != 0)                                //其他应付款
                    {
                        n = Convert.ToInt16(Kemu[8, 1]) - 1;
                        str = CU.Zifu(YEB[n, 1]);
                        List<string> Mingcheng = new List<string>();
                        List<double> Jiefang = new List<double>();
                        List<double> Daifang = new List<double>();

                        do
                        {
                            if ((n == N) || (!CU.Zifu(YEB[n + 1, 1]).Contains(CU.Zifu(YEB[n, 1]))))
                            {
                                Mingcheng.Add(YEB[n, 2].ToString().Trim());
                                Jiefang.Add(CU.Shuzi(YEB[n, 7]));
                                Daifang.Add(CU.Shuzi(YEB[n, 8]));
                            }
                            n = n + 1;
                        } while (CU.Zifu(YEB[n, 1]).Contains(str));
                        string[,] mingcheng = new string[Mingcheng.Count, 1];
                        double[,] jiefang = new double[Jiefang.Count, 1];
                        double[,] daifang = new double[Daifang.Count, 1];

                        int k = 0;
                        foreach (string s in Mingcheng)
                        {
                            mingcheng[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Jiefang)
                        {
                            jiefang[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Daifang)
                        {
                            daifang[k, 0] = s;
                            k++;
                        }

                        WorkingPaper.Wb.Sheets["其他应付"].Range["A14:A" + (13 + Mingcheng.Count).ToString()].Value2 = mingcheng;
                        WorkingPaper.Wb.Sheets["其他应付"].Range["B14:B" + (13 + Jiefang.Count).ToString()].Value2 = jiefang;
                        WorkingPaper.Wb.Sheets["其他应付"].Range["C14:C" + (13 + Daifang.Count).ToString()].Value2 = daifang;
                    }


                    if (CU.Shuzi(Kemu[11, 1]) != 0)                                //实收公积
                    {
                        n = Convert.ToInt16(Kemu[11, 1]) - 1;
                        str = CU.Zifu(YEB[n, 1]);
                        int i = 8;
                        double m = 0, m1, m2, m3;
                        m1 = Math.Round(CU.Shuzi(YEB[n, 4]) - CU.Shuzi(YEB[n, 3]), 2);//期初数
                        m2 = Math.Round(CU.Shuzi(YEB[n, 6]), 2);//本期增加
                        m3 = Math.Round(CU.Shuzi(YEB[n, 5]), 2);//本期减少
                        do
                        {
                            if ((n == N) || (!CU.Zifu(YEB[n + 1, 1]).Contains(CU.Zifu(YEB[n, 1]))))
                            {
                                m = Math.Round(CU.Shuzi(YEB[n, 7]) - CU.Shuzi(YEB[n, 8]), 2);
                                WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 2].Value = YEB[n, 2].ToString().Trim();
                                WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 3].Value = Math.Round(CU.Shuzi(YEB[n, 4]) - CU.Shuzi(YEB[n, 3]), 2);
                                WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 4].Value = Math.Round(CU.Shuzi(YEB[n, 6]), 2);
                                WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 5].Value = Math.Round(CU.Shuzi(YEB[n, 5]), 2);
                                i = i + 1;
                            }
                            n = n + 1;
                        } while (CU.Zifu(YEB[n, 1]).Contains(str) && i < 18);
                        if (i == 18)
                        {
                            WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 2].Value = "其他股东";
                            WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 3].Value = m1 - (double)WorkingPaper.Wb.Worksheets["实收公积"].Cells[13, 3].Value;
                            WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 4].Value = m2 - (double)WorkingPaper.Wb.Worksheets["实收公积"].Cells[13, 4].Value;
                            WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 5].Value = m3 - (double)WorkingPaper.Wb.Worksheets["实收公积"].Cells[13, 5].Value;
                        }
                    }

                    if (CU.Shuzi(Kemu[13, 1]) != 0)                                //盈余公积
                    {
                        n = Convert.ToInt16(Kemu[13, 1]) - 1;
                        str = CU.Zifu(YEB[n, 1]);
                        int i = 25;
                        double m = 0, m1, m2, m3;
                        m1 = Math.Round(CU.Shuzi(YEB[n, 4]) - CU.Shuzi(YEB[n, 3]), 2);//期初数
                        m2 = Math.Round(CU.Shuzi(YEB[n, 6]), 2);//本期增加
                        m3 = Math.Round(CU.Shuzi(YEB[n, 5]), 2);//本期减少
                        do
                        {
                            if ((n == N) || (!CU.Zifu(YEB[n + 1, 1]).Contains(CU.Zifu(YEB[n, 1]))))
                            {
                                m = Math.Round(CU.Shuzi(YEB[n, 7]) - CU.Shuzi(YEB[n, 8]), 2);
                                WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 2].Value = YEB[n, 2].ToString().Trim();
                                WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 3].Value = Math.Round(CU.Shuzi(YEB[n, 4]) - CU.Shuzi(YEB[n, 3]), 2);
                                WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 4].Value = Math.Round(CU.Shuzi(YEB[n, 6]), 2);
                                WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 5].Value = Math.Round(CU.Shuzi(YEB[n, 5]), 2);
                                i = i + 1;
                            }
                            n = n + 1;
                        } while (CU.Zifu(YEB[n, 1]).Contains(str) && i < 29);
                        if (i == 29)
                        {
                            WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 2].Value = "其他盈余公积";
                            WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 3].Value = m1 - (double)WorkingPaper.Wb.Worksheets["实收公积"].Cells[24, 3].Value;
                            WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 4].Value = m2 - (double)WorkingPaper.Wb.Worksheets["实收公积"].Cells[24, 4].Value;
                            WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 5].Value = m3 - (double)WorkingPaper.Wb.Worksheets["实收公积"].Cells[24, 5].Value;
                        }
                    }
                    WorkingPaper.Wb.Application.ScreenUpdating = true;
                    MessageBox.Show("填写底稿完成");
                }
                catch (Exception ex)
                {
                    Globals.WPToolAddln.Application.ScreenUpdating = true;
                    MessageBox.Show("用户操作出现错误：" + ex.Message);
                }
            }
        }

        public static void 税费填写()
        {
            if (WorkingPaper.OOO)
            {
                if (true)
                {
                    try
                    {
                        WorkingPaper.Wb.Application.ScreenUpdating = false;
                        Worksheet SH = WorkingPaper.Wb.Sheets["税金申报明细"];
                        //SH.Range["A:I"].Replace(SH.Range["R5"].Value, ""); //【税费缴纳测算】表
                        double jj = 0;
                        double[,] c = new double[13, 1], e = new double[13, 1], m = new double[6, 1], k = new double[6, 1];
                        int n = SH.Cells[SH.UsedRange.Rows.Count + 1, 1].End[XlDirection.xlUp].Row;
                        object[,] Shuifei = SH.Range["A2:L" + n.ToString()].Value2;
                        string year = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["B4"].Value2);
                        for (int i = 1; i <= n - 1; i++)
                        {
                            if (Shuifei[i, 2] != null && CU.Zifu(Shuifei[i, 5]).Substring(0, 4) == year)
                            {
                                string 征收项目 = CU.Zifu(Shuifei[i, 2]);
                                if (征收项目.Contains("印花税"))
                                {
                                    string 征收品目 = CU.Zifu(Shuifei[i, 3]);
                                    switch (征收品目)
                                    {
                                        case "购销合同":
                                            c[0, 0] = c[0, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[0, 0] = e[0, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "建筑安装工程承包合同":
                                            c[1, 0] = c[1, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[1, 0] = e[1, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "技术合同":
                                            c[2, 0] = c[2, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[2, 0] = e[2, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "财产租赁合同":
                                            c[3, 0] = c[3, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[3, 0] = e[3, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "仓储保管合同":
                                            c[4, 0] = c[4, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[4, 0] = e[4, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "财产保险合同":
                                            c[5, 0] = c[5, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[5, 0] = e[5, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "货物运输合同(按运输费用万分之五贴花)":
                                            c[6, 0] = c[6, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[6, 0] = e[6, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "加工承揽合同":
                                            c[7, 0] = c[7, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[7, 0] = e[7, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "建设工程勘察设计合同":
                                            c[8, 0] = c[8, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[8, 0] = e[8, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "产权转移书据":
                                            c[9, 0] = c[9, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[9, 0] = e[9, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "借款合同":
                                            c[10, 0] = c[10, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[10, 0] = e[10, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "其他营业账簿":
                                        case "权利、许可证照":
                                        case "其他凭证":
                                            c[11, 0] = c[11, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[11, 0] = e[11, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "资金账簿":
                                            c[12, 0] = c[12, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[12, 0] = e[12, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        default:
                                            break;
                                    }
                                }
                                else
                                {
                                    if (征收项目.Contains("消费税"))
                                    {
                                        m[0, 0] = m[0, 0] + CU.Shuzi(Shuifei[i, 11]);
                                    }
                                    else
                                    {
                                        if (征收项目.Contains("城市维护建设税"))
                                        {
                                            m[1, 0] = m[1, 0] + CU.Shuzi(Shuifei[i, 11]);
                                        }
                                        else
                                        {
                                            if (征收项目.Contains("教育费附加"))
                                            {
                                                m[2, 0] = m[2, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            }
                                            else
                                            {
                                                if (征收项目.Contains("地方教育附加"))
                                                {
                                                    m[3, 0] = m[3, 0] + CU.Shuzi(Shuifei[i, 11]);
                                                }
                                                else
                                                {
                                                    if (征收项目.Contains("资源税"))
                                                    {
                                                        m[4, 0] = m[4, 0] + CU.Shuzi(Shuifei[i, 11]);
                                                    }
                                                    else
                                                    {
                                                        if (征收项目.Contains("土地增值税"))
                                                        {
                                                            m[5, 0] = m[5, 0] + CU.Shuzi(Shuifei[i, 11]);
                                                        }
                                                        else
                                                        {
                                                                if (征收项目.Contains("房产税"))
                                                                {
                                                                    jj = jj + CU.Shuzi(Shuifei[i, 11]);
                                                                }
                                                                else
                                                                {
                                                                    if (征收项目.Contains("车船税"))
                                                                    {
                                                                        k[0, 0] = k[0, 0] + CU.Shuzi(Shuifei[i, 11]);
                                                                    }
                                                                    else
                                                                    {
                                                                        if (征收项目.Contains("城镇土地使用税"))
                                                                        {
                                                                            k[1, 0] = k[1, 0] + CU.Shuzi(Shuifei[i, 11]);
                                                                        }
                                                                        else
                                                                        {
                                                                            if (征收项目.Contains("契税"))
                                                                            {
                                                                                k[2, 0] = k[2, 0] + CU.Shuzi(Shuifei[i, 11]);
                                                                            }
                                                                            else
                                                                            {
                                                                                if (征收项目.Contains("残疾人就业保障金"))
                                                                                {
                                                                                    k[3, 0] = k[3, 0] + CU.Shuzi(Shuifei[i, 11]);
                                                                                }
                                                                                else
                                                                                {
                                                                                    if (征收项目.Contains("文化事业建设费"))
                                                                                    {
                                                                                        k[4, 0] = k[4, 0] + CU.Shuzi(Shuifei[i, 11]);
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        if (征收项目.Contains("个人所得税"))
                                                                                        {
                                                                                            k[5, 0] = k[5, 0] + CU.Shuzi(Shuifei[i, 11]);
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        c[11, 0] = e[11, 0] / 5;

                        WorkingPaper.Wb.Sheets["税费缴纳测算"].Range["C38:C50"].Value2 = c;
                        WorkingPaper.Wb.Sheets["税费缴纳测算"].Range["E38:E50"].Value2 = e;
                        WorkingPaper.Wb.Sheets["税费缴纳测算"].Cells[17, 5].Value = e[0, 0] + e[1, 0] + e[2, 0] + e[3, 0] +
                            e[4, 0] + e[5, 0] + e[6, 0] + e[7, 0] + e[8, 0] + e[9, 0] + e[10, 0] + e[11, 0];

                        WorkingPaper.Wb.Sheets["税费缴纳测算"].Range["E9:E14"].Value2 = m;
                        WorkingPaper.Wb.Sheets["税费缴纳测算"].Cells[16, 5].Value = jj;
                        WorkingPaper.Wb.Sheets["税费缴纳测算"].Range["E18:E23"].Value2 = k;
                        WorkingPaper.Wb.Application.ScreenUpdating = true;
                        MessageBox.Show("税费填写完成");
                    }
                    catch (Exception ex)
                    {
                        Globals.WPToolAddln.Application.ScreenUpdating = true;
                        MessageBox.Show("用户操作出现错误：" + ex.Message);
                    }
                }
            }
        }
    }
}
