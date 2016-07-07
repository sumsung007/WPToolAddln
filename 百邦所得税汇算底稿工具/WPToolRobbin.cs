using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Core;
using System.Text.RegularExpressions;

namespace 百邦所得税汇算底稿工具
{
    public partial class WorkingPaper
    {
        //
        Microsoft.Office.Tools.CustomTaskPane MyTaskpane;
        public static Workbook Wb;
        public static Boolean OOO=false;
        Contents Con;
        CommandBarButton Cd;
        //

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Con = new Contents();
            MyTaskpane = Globals.WPToolAddln.CustomTaskPanes.Add(Con, "税审底稿工具");
            MyTaskpane.Width = 300;
            MyTaskpane.VisibleChanged+=new EventHandler(MyTaskpane_VisibleChanged);
            if (! CU.授权检测())
            {
                tb显示目录.Enabled = false;
                btn基本情况.Enabled = false;
                btn余额报表.Enabled = false;
                btn税费测算.Enabled = false;
                btn检查表.Enabled = false;
                btn底稿打印.Enabled = false;
                btn底稿查看.Enabled = false;
                btn客户沟通.Enabled = false;
                btn查看报告.Enabled = false;
                btn导出报告.Enabled = false;
                btn工具设置.Enabled = false;
                btn注册.Visible = true;
                MessageBox.Show("底稿工具尚未注册，请进入设置后将机器码发给授权单位授权！");
            }
            Globals.WPToolAddln.Application.WorkbookActivate += Application_WorkbookActivate;
            
        }

        private void Application_SheetActivate(object Sh)
        {
            if(OOO)
            {
                string ss = Wb.ActiveSheet.Name;

                switch (ss)
                {
                    case "余额表":
                    case "税金申报明细":
                        Con.显示选项卡(ss);
                        break;
                    case "检查表":
                        Globals.WPToolAddln.Application.SheetFollowHyperlink += Application_SheetFollowHyperlink;
                        break;
                    case "A000000企业基础信息表":
                        Globals.WPToolAddln.Application.SheetSelectionChange += Application_SheetSelectionChange; 
                        break;
                    default:
                        Globals.WPToolAddln.Application.SheetFollowHyperlink -= Application_SheetFollowHyperlink;
                        Globals.WPToolAddln.Application.SheetSelectionChange -= Application_SheetSelectionChange;
                        Con.显示选项卡("");
                        break;
                }
            }
        }

        private void Application_SheetSelectionChange(object Sh, Range Target)
        {
            if (Target.Address == "$B$15:$F$15")
            {
                存货计价 ch= new 存货计价();
                
                ch.ShowDialog();
            }
        }
        

        private void Application_SheetFollowHyperlink(object Sh, Hyperlink Target)
        {
            if (WorkingPaper.OOO)
            {
                try
                {
                    Globals.WPToolAddln.Application.ScreenUpdating = false;
                    Wb.Worksheets[Target.Range.Value].Visible = true;
                    if (Wb.ActiveSheet.Cells[1, 7].Value.ToString() == "跳转超链接所选页面")
                        Wb.Worksheets[Target.Range.Value].Select();
                    Globals.WPToolAddln.Application.ScreenUpdating = true;
                }
                catch (Exception ex)
                {
                    Globals.WPToolAddln.Application.ScreenUpdating = true;
                    MessageBox.Show("用户操作出现错误：" + ex.Message);
                }
            }
        }
        private void 添加右键()
        {
            if (OOO)
            {
                if (Cd == null)
                {
                    Wb.Application.CommandBars["cell"].Reset();
                    //Wb.Application.CommandBars["column"].Reset();
                    Cd = (CommandBarButton)Wb.Application.CommandBars["cell"].Controls.Add(MsoControlType.msoControlButton,
                        1, Before: 1);
                    Cd.Caption = "返回首页";
                    //Cd.Picture=(stdole.IPictureDisp)百邦所得税汇算底稿工具.Properties.Resources.border;
                    Cd.Click += Cd_Click;
                }
            }
            else
            {
                Globals.WPToolAddln.Application.CommandBars["cell"].Reset();
                //Globals.WPToolAddln.Application.CommandBars["column"].Reset();
                Cd = null;
            }
        }

        private void Cd_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (WorkingPaper.OOO)
            {
                try
                {
                    WorkingPaper.Wb.Application.ScreenUpdating = false;
                    WorkingPaper.Wb.Worksheets[1].Visible = XlSheetVisibility.xlSheetVisible;
                    int C = WorkingPaper.Wb.Worksheets.Count;
                    for (int i = 2; i <= C; i++)
                    {
                        WorkingPaper.Wb.Worksheets[i].Visible = XlSheetVisibility.xlSheetVisible;
                    }
                    WorkingPaper.Wb.Worksheets[1].Visible = XlSheetVisibility.xlSheetHidden;
                    WorkingPaper.Wb.Sheets["主页"].Select();
                    WorkingPaper.Wb.Application.ScreenUpdating = true;
                }
                catch (Exception ex)
                {
                    Globals.WPToolAddln.Application.ScreenUpdating = true;
                    MessageBox.Show("用户操作出现错误：" + ex.Message);
                }
            }
        }

        private void Application_WorkbookActivate(Workbook Wb)
        {
            if (CU.文件判断())
            {
                MyTaskpane.Visible = true;
                Globals.WPToolAddln.Application.SheetActivate += Application_SheetActivate;
                string ss = Wb.ActiveSheet.Name;
                switch (ss)
                {
                    case "余额表":
                    case "税金申报明细":
                        Con.显示选项卡(ss);
                        break;
                    case "检查表":
                        Globals.WPToolAddln.Application.SheetFollowHyperlink += Application_SheetFollowHyperlink;
                        break;
                    case "A000000企业基础信息表":
                        Globals.WPToolAddln.Application.SheetSelectionChange += Application_SheetSelectionChange;
                        break;
                    default:
                        Globals.WPToolAddln.Application.SheetFollowHyperlink -= Application_SheetFollowHyperlink;
                        Globals.WPToolAddln.Application.SheetSelectionChange -= Application_SheetSelectionChange;
                        Con.显示选项卡("");
                        break;
                }
            }
            else
            {
                if (MyTaskpane != null) MyTaskpane.Visible = false;
                Globals.WPToolAddln.Application.SheetActivate -= Application_SheetActivate;
                Globals.WPToolAddln.Application.SheetFollowHyperlink -= Application_SheetFollowHyperlink;
            }
            添加右键();
        }

        private void MyTaskpane_VisibleChanged(object sender, EventArgs e)
        {
            tb显示目录.Checked=MyTaskpane.Visible;
        }

        private void btnHelp_Click(object sender, RibbonControlEventArgs e)         //关于程序
        {
            AboutBox1 AB = new AboutBox1();
            AB.ShowDialog();
            
        }

        private void btn新建_Click(object sender, RibbonControlEventArgs e)          //新建底稿文件
        {
            
            DialogResult dr = MessageBox.Show("是否新建一个税审底稿？", "新建", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                SaveFileDialog Sv = new SaveFileDialog();
                Sv.Filter = "百邦税审底稿(*.xlsx)|*.xlsx";
                Sv.FileName = "税审2015年底稿";
                Sv.Title = "保存新的税审底稿";
                Sv.OverwritePrompt = true;
                Sv.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
                if (Sv.ShowDialog() == DialogResult.OK)
                {
                    File.Copy(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "\\税审底稿模板.xlsx", Sv.FileName.ToString(), true);
                    Globals.WPToolAddln.Application.Workbooks.Open(Sv.FileName.ToString());
                }

                //Workbook wb = Globals.WPToolAddln.Application.Workbooks.Add(AppDomain.CurrentDomain.SetupInformation.ApplicationBase +"\\税审底稿模板.xltx");
                /*if (wb.Worksheets.Count>0)
                {
                    string[,] str = new string[,] { {"纳税人(委托人)税务登记号"}, { "纳税人(委托人)名称" }, { "鉴证审核年度" }, 
                        { "进行鉴证审核时间起" }, { "进行鉴证审核时间止" }, { "鉴证报告编号" }, { "鉴证报告意见" }, 
                        { "签名注册税务师1身份证号" }, {"签名注册税务师1姓名" }, { "签名注册税务师2身份证号" }, 
                        { "签名注册税务师2姓名" }, { "事务所税务登记号" }, { "事务所名称" } };
                    wb.Worksheets[1].Name = "About";
                    wb.Worksheets[1].Activate();
                    wb.ActiveSheet.range["A1:A13"].value = str;
                    wb.ActiveSheet.range["B13"].value = "厦门百邦税务师事务所有限公司";
                    wb.ActiveSheet.range["A1:B13"].EntireColumn.AutoFit();
                    wb.ActiveSheet.Protect("BaiBang12345");
                }*/
            }

        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            if(MyTaskpane ==null)
            {
                Con = new Contents();
                MyTaskpane = Globals.WPToolAddln.CustomTaskPanes.Add(Con, "税审底稿工具");
                MyTaskpane.Width = 300;
                MyTaskpane.VisibleChanged += new EventHandler(MyTaskpane_VisibleChanged);
            }
            MyTaskpane.Visible = tb显示目录.Checked;
        }

        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            if (WorkingPaper.OOO)
            {
                CU.工作表切换(new string[] { "A100000中华人民共和国企业所得税年度纳税申报表（A类）" ,
                "A000000企业基础信息表","A106000企业所得税弥补亏损明细表" ,"事项说明","凭证检查",
                "(二)附表-纳税调整额的审核","交换意见","当局声明" ,"业务约定","现金证明"});
                CU.事项说明();
            }
        }



        //菜单按键
        private void btn基本情况_Click(object sender, RibbonControlEventArgs e)
        {
            CU.工作表切换(new string[] { "地税、基本情况", "A000000企业基础信息表", "基本情况" });
        }

        private void btn余额报表_Click(object sender, RibbonControlEventArgs e)
        {
            CU.工作表切换(new string[] { "余额表", "资产负债", "利润" });
        }

        private void btn税费测算_Click(object sender, RibbonControlEventArgs e)
        {
            CU.工作表切换(new string[] { "纳税申报数据","主营税金","税费缴纳测算","纳税申报数据",
                "收入与申报核对表","企业各税审核汇总表","税金申报明细","社保明细工资人数","补亏","税金申报明细"});
        }

        private void btn检查表_Click(object sender, RibbonControlEventArgs e)
        {
            CU.工作表切换(new string[] { "凭证检查","检查表"});
            Wb.Sheets["检查表"].Rows["2:69"].Hidden = false;
            string s = "";
            double k;
            object[,] JCB = Wb.Sheets["检查表"].Range["C2:C69"].Value2;
            for (int i=1;i<=68;i++)
            {
                if (JCB[i, 1] != null)
                {
                    if (double.TryParse(JCB[i, 1].ToString().Trim(), out k))
                    {
                        
                        if (k == 0)
                        {
                            s = s + ",C" + (i+1).ToString();
                        }
                    }
                }
            }
            if(s.Length>0)
            {
                Wb.Sheets["检查表"].Range[s.Substring(1, s.Length - 1)].EntireRow.Hidden = true;
            }
        }

        public static void 报表填写()
        {
            if (CU.文件判断())
            {

                if ((Wb.Sheets["余额表"].Range["A2"] != null) && (Globals.WPToolAddln.Application.ActiveSheet.Name == "余额表") &&
                    (MessageBox.Show("是否填写报表？", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes))
                {
                    if (Wb.Sheets["基本情况"].Cells[8, 2].Value == "中汇百邦（厦门）税务师事务所有限公司" &&
                        Wb.Sheets["档案封面"].Cells[6, 1].Value == "中汇百邦（厦门）税务师事务所有限公司" &&
                        Wb.Sheets["基本情况（封面）"].Cells[16, 2].Value == "中汇百邦（厦门）税务师事务所有限公司")
                    {
                        try
                        {
                            string kemu, daima;
                            Wb.Application.ScreenUpdating = false;
                            Wb.Sheets["资产负债"].Range["C5:D18,C21:D24,C26:D33,G5:H14,G21:H24,G30:H33"].Value = "";
                            Wb.Sheets["利润"].Range["C5:D26,C28:D36,C38:D38,G6:H6,H7,G8:H10,G12:H18,G20:H21"].Value = "";
                            Worksheet SH = Wb.Sheets["余额表"];
                            int N = SH.Cells[SH.UsedRange.Rows.Count + 1, 2].End[XlDirection.xlUp].Row;
                            int changdu = (int)Wb.Sheets["余额表"].Range["o2"].Value;
                            object[,] YEB = SH.Range["A2:H" + N.ToString()].Value2;
                            double[] qc = new double[60], qm = new double[60], lrb = new double[14];

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
                                                        qc[59] = -Math.Round(CU.Shuzi(YEB[i, 3]) - CU.Shuzi(YEB[i, 4]), 2);
                                                        qm[59] = -Math.Round(CU.Shuzi(YEB[i, 7]) - CU.Shuzi(YEB[i, 8]), 2);
                                                    }
                                                }
                                            }
                                            break;
                                    }
                                }
                            }
                            double[,] ldzc = new double[14, 2];
                            ldzc[0, 0] = qc[0] + qc[1] + qc[2];             //货币资金
                            ldzc[0, 1] = qm[0] + qm[1] + qm[2];
                            ldzc[1, 0] = qc[3] + qc[4];                     //短期投资
                            ldzc[1, 1] = qm[3] + qm[4];
                            ldzc[2, 0] = qc[5];                             //应收票据
                            ldzc[2, 1] = qm[5];
                            ldzc[3, 0] = qc[6];                             //应收账款
                            ldzc[3, 1] = qm[6];
                            ldzc[4, 0] = qc[7];                             //预付账款
                            ldzc[4, 1] = qm[7];
                            ldzc[5, 0] = qc[8];                             //应收股息
                            ldzc[5, 1] = qm[8];
                            ldzc[6, 0] = qc[9];                             //应收利息
                            ldzc[6, 1] = qm[9];
                            ldzc[7, 0] = qc[10];                             //其他应收款
                            ldzc[7, 1] = qm[10];
                            ldzc[8, 0] = qc[11] + qc[12] + qc[13] + qc[14] + qc[15] + qc[16] + qc[17] + qc[18] +
                                qc[19] + qc[20] + qc[21] + qc[22] + qc[57] + qc[58];    //存货
                            ldzc[8, 1] = qm[11] + qm[12] + qm[13] + qm[14] + qm[15] + qm[16] + qm[17] + qm[18] +
                                qm[19] + qm[20] + qm[21] + qm[22] + qm[57] + qm[58];
                            ldzc[13, 0] = qc[23] + qc[24];    //其他流动资产=待摊费用+应收出口退税
                            ldzc[13, 1] = qm[23] + qm[24];


                            Wb.Sheets["资产负债"].Range["C5:D18"].Value2 = ldzc;

                            ldzc = new double[4, 2];
                            ldzc[0, 0] = qc[25];//长期债权投资
                            ldzc[0, 1] = qm[25];
                            ldzc[1, 0] = qc[26] + qc[27] + qc[28];//长期股权投资
                            ldzc[1, 1] = qm[26] + qm[27] + qm[28];
                            ldzc[2, 0] = qc[29];//固定资产
                            ldzc[2, 1] = qm[29];
                            ldzc[3, 0] = qc[30];//累计折旧
                            ldzc[3, 1] = qm[30];
                            Wb.Sheets["资产负债"].Range["C21:D24"].Value2 = ldzc;

                            ldzc = new double[7, 2];
                            ldzc[0, 0] = qc[32];                             //在建工程
                            ldzc[0, 1] = qm[32];
                            ldzc[1, 0] = qc[31];                             //工程物资
                            ldzc[1, 1] = qm[31];
                            ldzc[2, 0] = qc[33];                             //固定资产清理
                            ldzc[2, 1] = qm[33];
                            ldzc[3, 0] = qc[34];                             //生产性生物资产
                            ldzc[3, 1] = qm[34];
                            ldzc[4, 0] = qc[35];                             //无形资产
                            ldzc[4, 1] = qm[35];
                            ldzc[5, 0] = qc[36];                             //开发支出
                            ldzc[5, 1] = qm[36];
                            ldzc[6, 0] = qc[37];                             //长期待摊费用
                            ldzc[6, 1] = qm[37];
                            Wb.Sheets["资产负债"].Range["C26:D32"].Value2 = ldzc;

                            ldzc = new double[10, 2];
                            ldzc[0, 0] = qc[38];                             //短期借款
                            ldzc[0, 1] = qm[38];
                            ldzc[1, 0] = qc[39];                             //应付票据
                            ldzc[1, 1] = qm[39];
                            ldzc[2, 0] = qc[40];                             //应付账款
                            ldzc[2, 1] = qm[40];
                            ldzc[3, 0] = qc[41];                             //预收账款
                            ldzc[3, 1] = qm[41];
                            ldzc[4, 0] = qc[42] + qc[43];                     //应付职工薪酬=应付工资+应付福利费
                            ldzc[4, 1] = qm[42] + qm[43];
                            ldzc[5, 0] = qc[45] + qc[46];                     //应交税费=应交税金+其他应交款
                            ldzc[5, 1] = qm[45] + qm[46];
                            ldzc[6, 0] = qc[47];                             //应付利息
                            ldzc[6, 1] = qm[47];
                            ldzc[7, 0] = qc[44];                             //应付利润
                            ldzc[7, 1] = qm[44];
                            ldzc[8, 0] = qc[48];                             //其他应付款
                            ldzc[8, 1] = qm[48];
                            ldzc[9, 0] = qc[49];                             //其他流动负债 含预提费用
                            ldzc[9, 1] = qm[49];
                            Wb.Sheets["资产负债"].Range["G5:H14"].Value2 = ldzc;


                            ldzc = new double[2, 2];
                            ldzc[0, 0] = qc[50];//长期借款
                            ldzc[0, 1] = qm[50];
                            ldzc[1, 0] = qc[51];//长期应付款
                            ldzc[1, 1] = qm[51];

                            Wb.Sheets["资产负债"].Range["G21:H22"].Value2 = ldzc;

                            ldzc = new double[4, 2];
                            ldzc[0, 0] = qc[59];//实收资本
                            ldzc[0, 1] = qm[59];
                            ldzc[1, 0] = qc[52];//资本公积
                            ldzc[1, 1] = qm[52];
                            ldzc[2, 0] = qc[53];//盈余公积
                            ldzc[2, 1] = qm[53];
                            ldzc[3, 0] = qc[54] + qc[55];//未分配利润+本年利润
                            ldzc[3, 1] = qm[54] + qm[55];

                            Wb.Sheets["资产负债"].Range["G30:H33"].Value2 = ldzc;

                            ldzc = new double[22, 1];
                            ldzc[0, 0] = lrb[0] + lrb[3];//营业收入=主营业务收入+其他业务收入
                            ldzc[1, 0] = lrb[1] + lrb[4];//营业成本=主营业务成本+其他业务成本
                            ldzc[2, 0] = lrb[2];//营业税金及附加
                            ldzc[10, 0] = lrb[5];//销售费用
                            ldzc[13, 0] = lrb[6];//管理费用
                            ldzc[17, 0] = lrb[7];//财务费用
                            ldzc[19, 0] = lrb[12];//资产减值损失
                            ldzc[20, 0] = lrb[13];//公允价值变动损益
                            ldzc[21, 0] = lrb[8];//投资收益
                            Wb.Sheets["利润"].Range["C5:C26"].Value2 = ldzc;

                            ldzc = new double[3, 1];
                            ldzc[0, 0] = lrb[9];//营业外收入
                            ldzc[2, 0] = lrb[10];//营业外支出
                            Wb.Sheets["利润"].Range["C28:C30"].Value2 = ldzc;

                            Wb.Sheets["利润"].Cells[38, 3].Value = lrb[11];//所得税
                            WorkingPaper.Wb.Application.ScreenUpdating = true;
                        }
                        catch (Exception ex)
                        {
                            WorkingPaper.Wb.Application.ScreenUpdating = true;
                            MessageBox.Show("用户操作出现错误：" + ex.Message);
                        }
                        if (CU.Shuzi(Wb.Sheets["资产负债"].Range["H38"].Value2) != 0)
                        {
                            MessageBox.Show("报表填写完毕，请复查!" + "资产负债表未分配利润期末余额与利润表期末未分配利润差异" +
                                CU.Shuzi(Wb.Sheets["资产负债"].Range["H38"].Value2).ToString("N"));
                        }
                        else
                            MessageBox.Show("报表填写完毕，请复查!");
                    }
                }
            }
        }

        private void btn注册_Click(object sender, RibbonControlEventArgs e)
        {
            REGForm reg = new REGForm();
            if(reg.ShowDialog()==DialogResult.Yes)
            {
                tb显示目录.Enabled = true;
                btn基本情况.Enabled = true;
                btn余额报表.Enabled = true;
                btn税费测算.Enabled = true;
                btn检查表.Enabled = true;
                btn底稿打印.Enabled = true;
                btn底稿查看.Enabled = true;
                btn客户沟通.Enabled = true;
                btn查看报告.Enabled = true;
                btn导出报告.Enabled = true;
                btn工具设置.Enabled = true;
                btn注册.Visible = false;
            }
        }

        private void btn底稿查看_Click(object sender, RibbonControlEventArgs e)
        {
            if (WorkingPaper.OOO)
            {
                try
                {
                    WorkingPaper.Wb.Application.ScreenUpdating = false;
                    WorkingPaper.Wb.Worksheets[1].Visible = XlSheetVisibility.xlSheetVisible;
                    int C = WorkingPaper.Wb.Worksheets.Count;
                    for (int i = 2; i <= C; i++)
                    {
                        WorkingPaper.Wb.Worksheets[i].Visible = XlSheetVisibility.xlSheetVisible;
                    }
                    WorkingPaper.Wb.Worksheets[1].Visible = XlSheetVisibility.xlSheetHidden;
                    WorkingPaper.Wb.Sheets["主页"].Select();
                    WorkingPaper.Wb.Application.ScreenUpdating = true;
                }
                catch (Exception ex)
                {
                    WorkingPaper.Wb.Application.ScreenUpdating = true;
                    MessageBox.Show("用户操作出现错误：" + ex.Message);
                }
            }
        }

        void 查看报告()
        {
            CU.工作表切换(new string[] { "基本情况（封面）", "1.保留意见", "2.否定意见", "3.无保留意见", "4.无法表明意见", "(二)企业基本情况和审核事项说明", "(二)附表-科目说明",
                "(二)附表-纳税调整额的审核", "（三）企业所得税年度纳税申报表填报表单", "A000000企业基础信息表", "A100000中华人民共和国企业所得税年度纳税申报表（A类）", "A101010一般企业收入明细表",
                "A101020金融企业收入明细表", "A102010一般企业成本支出明细表", "A102020金融企业支出明细表", "A103000事业单位、民间非营利组织收入、支出明细表", "A104000期间费用明细表",
                "A105000纳税调整项目明细表", "A105010视同销售和房地产开发企业特定业务纳税调整明细表", "A105020未按权责发生制确认收入纳税调整明细表", "A105030投资收益纳税调整明细表",
                "A105040专项用途财政性资金纳税调整表", "A105050职工薪酬纳税调整明细表", "A105060广告费和业务宣传费跨年度纳税调整明细表", "A105070捐赠支出纳税调整明细表", "A105080资产折旧、摊销情况及纳税调整明细表",
                "A105081固定资产加速折旧、扣除明细表", "A105090资产损失税前扣除及纳税调整明细表", "A105091资产损失（专项申报）税前扣除及纳税调整明细表", "A105100企业重组纳税调整明细表",
                "A105110政策性搬迁纳税调整明细表", "A105120特殊行业准备金纳税调整明细表", "A106000企业所得税弥补亏损明细表", "A107010免税、减计收入及加计扣除优惠明细表", "A107011股息红利优惠明细表",
                "A107012综合利用资源生产产品取得的收入优惠明细表", "A107013金融保险等机构取得涉农利息保费收入优惠明细表", "A107014研发费用加计扣除优惠明细表", "A107020所得减免优惠明细表",
                "A107030抵扣应纳税所得额明细表", "A107040减免所得税优惠明细表", "A107041高新技术企业优惠情况及明细表", "A107042软件、集成电路企业优惠情况及明细表", "A107050税额抵免优惠明细表",
                "A108000境外所得税收抵免明细表", "A108010境外所得纳税调整后所得明细表", "A108020境外分支机构弥补亏损明细表", "A108030跨年度结转抵免境外所得税明细表", "A109000跨地区经营汇总纳税企业年度分摊企业所得税明细表",
                "A109010企业所得税汇总纳税分支机构所得税分配表", "A110010特殊性处理报告表", "A110011债务重组报告表", "A110012股权收购报告表 ", "A110013资产收购报告表", "A110014企业合并报告表 ", "A110015企业分立申报表",
                "A110016非货币资产投资递延纳税调整表", "A110017居民企业资产（股权）划转特殊性税务处理申报表", "分支机构企业所得税申报表（A类）", "（四）企业各税（费）审核汇总表", "（五）社会保险费明细表" });
                WorkingPaper.Wb.Sheets["基本情况（封面）"].Select();
                CU.事项说明();
        }

        private void btn查看报告_Click(object sender, RibbonControlEventArgs e)
        {
            if (WorkingPaper.OOO)
            {
                try
                {
                    WorkingPaper.Wb.Application.ScreenUpdating = false;
                    查看报告();
                    WorkingPaper.Wb.Application.ScreenUpdating = true;
                }
                catch (Exception ex)
                {
                    WorkingPaper.Wb.Application.ScreenUpdating = true;
                    MessageBox.Show("用户操作出现错误：" + ex.Message);
                }
            }
        }


        private void splitButton1_Click(object sender, RibbonControlEventArgs e)
        {

            if (WorkingPaper.OOO)
            {
                CU.工作表切换(new string[] { "A100000中华人民共和国企业所得税年度纳税申报表（A类）" ,
                "A000000企业基础信息表","A106000企业所得税弥补亏损明细表" ,"事项说明","凭证检查",
                "(二)附表-纳税调整额的审核","交换意见","当局声明" ,"业务约定"});
                CU.事项说明();
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if (WorkingPaper.OOO)
            {
                if (MessageBox.Show("现在将当前可见工作表导出为03版本Excel。是否继续？", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    SaveFileDialog Sv = new SaveFileDialog();
                    Sv.Filter = "Excel 2003工作簿(*.xls)|*.xls";
                    Sv.FileName = "税审工作表导出";
                    Sv.Title = "导出当前可见工作表";
                    Sv.OverwritePrompt = true;
                    Sv.InitialDirectory = WorkingPaper.Wb.Path;
                    //Sv.RestoreDirectory = true;
                    if (Sv.ShowDialog() == DialogResult.OK)
                    {
                        //string LocalFilePath = Sv.FileName.ToString(); //获得文件路径
                        //string FileNameExt = LocalFilePath.Substring(LocalFilePath.LastIndexOf("\\") + 1); //获取文件名，不带路径
                        //string FilePath = LocalFilePath.Substring(0, LocalFilePath.LastIndexOf("\\"));//获取文件路径，不带文件名 
                        try
                        {
                            Globals.WPToolAddln.Application.StatusBar = "正在导出可见工作表...";
                            Globals.WPToolAddln.Application.DisplayAlerts = false;
                            Globals.WPToolAddln.Application.ScreenUpdating = false;
                            Workbook Oldbook = WorkingPaper.Wb;
                            List<string> lists = new List<string>();
                            foreach (Worksheet ssh in Oldbook.Worksheets)
                            {
                                if (ssh.Visible == XlSheetVisibility.xlSheetVisible)
                                {
                                    lists.Add(ssh.Name.ToString());
                                }
                            }
                            string[] s = lists.ToArray();
                            Workbook Newbook = Globals.WPToolAddln.Application.Workbooks.Add();
                            int C = Newbook.Worksheets.Count;
                            Oldbook.Worksheets[s].Copy(Type.Missing, Newbook.Worksheets[C]);
                            for (int i = 1; i <= C; i++)
                            {
                                Newbook.Worksheets[1].Delete();
                            }
                            foreach (Name nm in Newbook.Names)
                            {
                                if (Regex.IsMatch(nm.RefersTo.ToString(), @"(#REF!)|\/|\\|\*|\[|\]"))
                                {
                                    nm.Delete();
                                }
                            }
                            Newbook.BreakLink(Oldbook.Path + "\\" + Oldbook.Name, XlLinkType.xlLinkTypeExcelLinks);
                            Newbook.SaveAs(Sv.FileName.ToString(), XlFileFormat.xlExcel8);
                            Newbook.Close();
                            Newbook = null;
                            Globals.WPToolAddln.Application.DisplayAlerts = true;
                            Globals.WPToolAddln.Application.ScreenUpdating = true;
                            Globals.WPToolAddln.Application.StatusBar = false;
                            MessageBox.Show("文件导出完成！");
                        }
                        catch (Exception ex)
                        {
                            Globals.WPToolAddln.Application.DisplayAlerts = true;
                            Globals.WPToolAddln.Application.ScreenUpdating = true;
                            Globals.WPToolAddln.Application.StatusBar = false;
                            MessageBox.Show("用户操作出现错误：" + ex.Message);
                        }
                    }

                }
            }
        }

        private void 底稿升级_Click(object sender, RibbonControlEventArgs e)
        {
            if (WorkingPaper.OOO)
            {
                string Banben1 = CU.Zifu(WorkingPaper.Wb.Worksheets["首页"].Range["A1"].Value2);
                string Banben;
                bool 升级;
                switch  (Banben1)
                    {
                    case "V20160508":
                    case "V20160508-0504":
                    case "V20160508-0504-0316":
                        Banben = Banben1;
                        升级 = false;
                        break;
                    case "V20160504":
                        if (WorkingPaper.Wb.Worksheets["资产负债"].Range["I20"].Formula == "=SUM(I9:I19)")
                            Banben = "V20160504";
                        else
                            Banben = "V20160504-0316";
                        升级 = true;
                        break;
                    default:
                        Banben = "V20160316";
                        升级 = true;
                        break;
                    }
                
                if (升级)
                {
                    if (MessageBox.Show("当前版本为："+Banben+"，最新版本为：V20160508。是否升级？", "提示！",
                        MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        if (MessageBox.Show("本操作具有不稳定性，会先保存当前文件，并以BAK后缀文件备份在文件同目录下。是否继续？", "警告！",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                        {
                            Globals.WPToolAddln.Application.StatusBar = "正在升级底稿...";
                            string fullname = WorkingPaper.Wb.FullName;
                            string number = "";
                            int i = 0;
                            while (File.Exists(fullname + ".bak" + number))
                            {
                                i++;
                                number = i.ToString();
                            }
                            WorkingPaper.Wb.Save();
                            File.Copy(WorkingPaper.Wb.FullName, fullname + ".bak" + number, true);

                            if (Banben == "V20160316")
                            {
                                #region 0316升级为0504更新过程
                                //插入表格

                                #region 事务所更名
                                WorkingPaper.Wb.Worksheets["基本情况"].Cells[8, 2].Value2 = "中汇百邦（厦门）税务师事务所有限公司";
                                WorkingPaper.Wb.Worksheets["档案封面"].Cells[6, 1].Value2 = "中汇百邦（厦门）税务师事务所有限公司";
                                WorkingPaper.Wb.Worksheets["基本情况（封面）"].Cells[16, 2].Value2 = "中汇百邦（厦门）税务师事务所有限公司";
                                WorkingPaper.Wb.Worksheets["基本情况（封面）"].Cells[15, 2].Value2 = "91350200776046719Q";
                                WorkingPaper.Wb.Worksheets["基本情况（封面）"].Range["E83,E86"].NumberFormatLocal = "@";
                                WorkingPaper.Wb.Worksheets["基本情况（封面）"].Range["E83:E84,E86:E87"].Borders.LineStyle = XlLineStyle.xlContinuous;
                                WorkingPaper.Wb.Worksheets["基本情况（封面）"].Range["E83"].Value2 = "350784197902181021";
                                WorkingPaper.Wb.Worksheets["基本情况（封面）"].Range["E84"].Value2 = "叶瑞卿";
                                WorkingPaper.Wb.Worksheets["基本情况（封面）"].Range["E86"].Value2 = "350623198105204207";
                                WorkingPaper.Wb.Worksheets["基本情况（封面）"].Range["E87"].Value2 = "陈酉凤";
                                WorkingPaper.Wb.Worksheets["首页"].Range["A1"].Value2 = "V20160504-0316";
                                #endregion

                                #region 其他业务
                                WorkingPaper.Wb.Worksheets["其他业务"].Range["F7"].Value2 = "成本";
                                WorkingPaper.Wb.Worksheets["其他业务"].Range["F8,F11,F13"].Interior.Pattern = XlPattern.xlPatternNone;
                                WorkingPaper.Wb.Worksheets["其他业务"].Range["F8,F11,F13"].ClearContents();
                                WorkingPaper.Wb.Worksheets["其他业务"].Rows["21:21"].Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                                WorkingPaper.Wb.Worksheets["其他业务"].Range["C20:J20"].FormulaR1C1 =
                                    @"=IF(其他业务!R18C3+主营收支!R20C8<>利润!R5C3,""营业收入账载数与报表数相差""&DOLLAR(其他业务!R18C3+主营收支!R20C8-利润!R5C3,2)&""元！"",""营业收入账载数与报表数相符！"")";
                                WorkingPaper.Wb.Worksheets["其他业务"].Range["A20:B21"].UnMerge();
                                WorkingPaper.Wb.Worksheets["其他业务"].Range["A20:B21"].Merge();
                                WorkingPaper.Wb.Worksheets["其他业务"].Range["C21:J21"].Merge();
                                WorkingPaper.Wb.Worksheets["其他业务"].Range["C21:J21"].FormulaR1C1 =
                                    @"=IF(其他业务!R18C8+主营收支!R37C8<>利润!R6C3,""营业成本账载数与报表数相差""&DOLLAR(其他业务!R18C8+主营收支!R37C8-利润!R6C3,2)&""元！"",""营业成本账载数与报表数相符！"")";
                                WorkingPaper.Wb.Worksheets["其他业务"].Range["C20:J21"].Interior.Color = 12632256;
                                #endregion

                                #region 主营收支
                                WorkingPaper.Wb.Worksheets["主营收支"].Range["C41:H41"].FormulaR1C1 =
                                    @"=IF(其他业务!R18C3+主营收支!R20C8<>利润!R5C3,""营业收入账载数与报表数相差""&DOLLAR(其他业务!R18C3+主营收支!R20C8-利润!R5C3,2)&""元！"",""营业收入账载数与报表数相符！"")";
                                WorkingPaper.Wb.Worksheets["主营收支"].Range["C42:H42"].FormulaR1C1 =
                                    @"=IF(其他业务!R18C8+主营收支!R37C8<>利润!R6C3,""营业成本账载数与报表数相差""&DOLLAR(其他业务!R18C8+主营收支!R37C8-利润!R6C3,2)&""元！"",""营业成本账载数与报表数相符！"")";
                                #endregion

                                #region 更正“1-3年”名称
                                Name Nm = WorkingPaper.Wb.Names.Item(Index: "一至三年");
                                Nm.RefersToR1C1 = "=A106000企业所得税弥补亏损明细表!R6C6:R8C6";
                                #endregion

                                #region 减免所得税优惠审核表
                                WorkingPaper.Wb.Worksheets["减免所得税优惠审核表"].Range["F1"].Value2 = "实际经营期";
                                WorkingPaper.Wb.Worksheets["减免所得税优惠审核表"].Range["G1"].Value2 = "从业人数";
                                WorkingPaper.Wb.Worksheets["减免所得税优惠审核表"].Range["H1"].Value2 = "资产总额";
                                WorkingPaper.Wb.Worksheets["减免所得税优惠审核表"].Range["F2"].FormulaR1C1 = "=截止月-起始月+1";
                                WorkingPaper.Wb.Worksheets["减免所得税优惠审核表"].Range["G2"].FormulaR1C1 =
                                    @"=IF(RC[-1]=12,ROUND((社保明细工资人数!R[6]C[3]/2+社保明细工资人数!R[9]C[3]+社保明细工资人数!R[12]C[3]+社保明细工资人数!R[15]C[3]+社保明细工资人数!R[18]C[3]/2)/4,0),""请自行计算"")";
                                WorkingPaper.Wb.Worksheets["减免所得税优惠审核表"].Range["H2"].FormulaR1C1 =
                                    @"=IF(RC[-2]=12,ROUND((资产负债!R[33]C[-5]/2+资产负债!R[9]C[1]+资产负债!R[12]C[1]+资产负债!R[15]C[1]+资产负债!R[33]C[-4]/2)/4/10000,2),""请自行计算"")";
                                WorkingPaper.Wb.Worksheets["减免所得税优惠审核表"].Range["H3"].FormulaR1C1 =
                                    @"=IF(A000000企业基础信息表!R[4]C[-3]=""否"",IF(AND(LEFT(A000000企业基础信息表!R[4]C[-6],2)>""05"",LEFT(A000000企业基础信息表!R[4]C[-6],2)<""47"",R[-1]C[-1]<=100,R[-1]C<=30000000),""工业企业"",IF(OR(LEFT(A000000企业基础信息表!R[4]C[-6],2)>""50"",LEFT(A000000企业基础信息表!R[4]C[-6],2)<""06""),IF(AND(R[-1]C[-1]<=80,R[-1]C<=10000000),""其他企业"",""""),"""")),"""")";
                                /*WorkingPaper.Wb.Worksheets["减免所得税优惠审核表"].Range["C4"].FormulaR1C1 =
                                    @"=IF(R[-1]C[5]<>"""",IF('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!R[22]C[1]<=200000,ROUND('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!R[22]C[1]*0.15,2),IF('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!R[22]C[1]<=300000,ROUND('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!R[22]C[1]*R[-1]C[6],2))),"""")";
                                WorkingPaper.Wb.Worksheets["减免所得税优惠审核表"].Range["C5"].FormulaR1C1 =
                                    @"=IF(R[-2]C[5]<>"""",IF('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!R[21]C[1]<=200000,ROUND('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!R[21]C[1]*0.15,2),IF('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!R[21]C[1]<=300000,ROUND('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!R[21]C[1]*R[-2]C[7],2))),"""")";
                                    */
                                WorkingPaper.Wb.Worksheets["减免所得税优惠审核表"].Range["C4:C5"].Interior.Pattern = XlPattern.xlPatternNone;
                                WorkingPaper.Wb.Worksheets["减免所得税优惠审核表"].Range["C4:C5"].ClearContents();
                                WorkingPaper.Wb.Worksheets["减免所得税优惠审核表"].Range["H4"].FormulaR1C1 =
                                    @"=IF(R[-1]C<>"""",IF('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!R[22]C[-4]<=200000,ROUND('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!R[22]C[-4]*0.15,2),IF('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!R[22]C[-4]<=300000,ROUND('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!R[22]C[-4]*R[-1]C[1],2))),"""")";
                                WorkingPaper.Wb.Worksheets["减免所得税优惠审核表"].Range["H5"].FormulaR1C1 =
                                    @"=IF(R[-2]C<>"""",IF('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!R[21]C[-4]<=200000,ROUND('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!R[21]C[-4]*0.15,2),IF('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!R[21]C[-4]<=300000,ROUND('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!R[21]C[-4]*R[-2]C[2],2))),"""")";
                                WorkingPaper.Wb.Worksheets["减免所得税优惠审核表"].Range["H6"].Value2 = "确认符合小微企业条件时，将H4、H5单元格金额填在C4、C5。";
                                #endregion

                                #region 检查表
                                WorkingPaper.Wb.Worksheets["检查表"].Range["C25"].FormulaR1C1 =
                                    @"=IF(OR(主营收支!R20C8+其他业务!R18C3<>利润!R5C3,主营收支!R37C8+其他业务!R18C8<>利润!R6C3),""不符"",0)";
                                WorkingPaper.Wb.Worksheets["检查表"].Range["C33"].FormulaR1C1 =
                                    @"=IF(OR(主营收支!R20C8+其他业务!R18C3<>利润!R5C3,主营收支!R37C8+其他业务!R18C8<>利润!R6C3),""不符"",0)";
                                WorkingPaper.Wb.Worksheets["检查表"].Range["C43"].FormulaR1C1 =
                                    @"=IF(OR(待摊预提!R12C8<>资产负债!R18C4,待摊预提!R19C8<>资产负债!R14C8),""不符"",0)";
                                WorkingPaper.Wb.Worksheets["检查表"].Range["C44"].FormulaR1C1 = @"=在建工程审核表!R15C6-资产负债!R26C4";
                                WorkingPaper.Wb.Worksheets["检查表"].Range["C48"].FormulaR1C1 =
                                    @"=IF(OR(实收公积!R13C6<>资产负债!R30C8,实收公积!R18C6<>资产负债!R31C8,实收公积!R24C6<>资产负债!R32C8),""不符"",0)";
                                #endregion

                                #region 待摊预提
                                WorkingPaper.Wb.Worksheets["待摊预提"].Range["C20:F20"].FormulaR1C1 =
                                    @"=IF(R12C8<>资产负债!R18C4,""待摊费用账载数与报表数相差""&DOLLAR(R12C8-资产负债!R18C4,2)&""元！"",""待摊费用账载数与报表数相符！"")";
                                WorkingPaper.Wb.Worksheets["待摊预提"].Range["G20:J20"].FormulaR1C1 =
                                    @"=IF(R19C8<>资产负债!R14C8,""预提费用账载数与报表数相差""&DOLLAR(R19C8-资产负债!R14C8,2)&""元！"",""预提费用账载数与报表数相符！"")";
                                #endregion

                                #region 抵扣应纳税所得额审核表
                                WorkingPaper.Wb.Worksheets["抵扣应纳税所得额审核表"].Range["C10"].FormulaR1C1 =
                                    @"=MAX('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!R22C4-'A100000中华人民共和国企业所得税年度纳税申报表（A类）'!R23C4-R18C3,0)";
                                #endregion

                                #region A106000企业所得税弥补亏损明细表
                                WorkingPaper.Wb.Worksheets["A106000企业所得税弥补亏损明细表"].Range["J6"].FormulaR1C1 =
                                    @"=IF(补亏!R[10]C[-4]<=0,0,IF(补亏!R[6]C[-4]>=0,0,IF(补亏!R[10]C[-4]>-补亏!R[6]C[-4]-RC[-3]-RC[-2]-RC[-1],-补亏!R[6]C[-4]-RC[-3]-RC[-2]-RC[-1],补亏!R[10]C[-4])))";
                                //@"=IF(补亏!R[10]C[-4]<=0,0,IF(补亏!R[6]C[-4]>=0,0,IF(-SUMIF(补亏!R[7]C[-4]:R[10]C[-4],""<0"")-RC[-3]-RC[-2]-RC[-1]-R[1]C[-2]-R[1]C[-1]-R[2]C[-1]>0,IF(补亏!R[10]C[-4]>-补亏!R[6]C[-4]-RC[-3]-RC[-2]-RC[-1],-补亏!R[6]C[-4]-RC[-3]-RC[-2]-RC[-1],补亏!R[10]C[-4]),0)))";
                                #endregion

                                #region 基本情况（未成功）
                                //MessageBox.Show(WorkingPaper.Wb.Worksheets["基本情况"].Range["B48:E48"].Formula.ToString());
                                #endregion

                                #region 招待
                                WorkingPaper.Wb.Worksheets["招待"].Rows["25:25"].Delete(XlDeleteShiftDirection.xlShiftUp);
                                WorkingPaper.Wb.Worksheets["其他事项"].Range["E33"].FormulaR1C1 = @"=-(广宣!R[-1]C[5])";
                                #endregion

                                #region 主营税金
                                WorkingPaper.Wb.Worksheets["主营税金"].Range["C8:C11"].FormulaR1C1 = @"=收入与申报核对表!R[6]C[4]";
                                WorkingPaper.Wb.Worksheets["主营税金"].Range["C12:C13"].FormulaR1C1 = @"=收入与申报核对表!R[8]C[4]";
                                #endregion

                                #region 加速折旧
                                WorkingPaper.Wb.Worksheets["A105081固定资产加速折旧、扣除明细表"].Range["A1"].Value2 = "A105081";
                                WorkingPaper.Wb.Worksheets["A105081固定资产加速折旧、扣除明细表"].Rows["5:5"].Delete(XlDeleteShiftDirection.xlShiftUp);
                                #endregion

                                #region 基本情况（封面）
                                WorkingPaper.Wb.Worksheets["基本情况（封面）"].Range["A1:B1"].Value2 = "企业所得税汇算清缴纳税申报鉴证报告（其他企业）";
                                WorkingPaper.Wb.Worksheets["基本情况（封面）"].Range["C1"].Value2 = "SSJZ6101.2";
                                WorkingPaper.Wb.Worksheets["基本情况（封面）"].Range["E41"].Value2 = "350221197708092567";
                                #endregion

                                #region A105080资产折旧、摊销情况及纳税调整明细表
                                WorkingPaper.Wb.Worksheets["A105080资产折旧、摊销情况及纳税调整明细表"].Range["I6:I11"].Interior.Pattern = XlPattern.xlPatternNone;
                                WorkingPaper.Wb.Worksheets["A105080资产折旧、摊销情况及纳税调整明细表"].Range["I6:I11"].ClearContents();
                                #endregion

                                #region A000000企业基础信息表
                                WorkingPaper.Wb.Worksheets["A000000企业基础信息表"].Range["E21:E25"].FormulaR1C1 =
                                    @"=IFERROR(基本情况!R[31]C[-2],"""")";
                                WorkingPaper.Wb.Worksheets["A000000企业基础信息表"].Range["C7:D7"].Value2 = "107从事国家限制或禁止行业";
                                WorkingPaper.Wb.Worksheets["A000000企业基础信息表"].Range["E7:F7"].Value2 = "否";
                                WorkingPaper.Wb.Worksheets["A000000企业基础信息表"].Range["B16:F16"].Value2 = "直接核销法";
                                #endregion

                                #region A100000中华人民共和国企业所得税年度纳税申报表（A类）
                                WorkingPaper.Wb.Worksheets["A100000中华人民共和国企业所得税年度纳税申报表（A类）"].Range["D24"].FormulaR1C1 =
                                    @"=MAX(MIN(R[-2]C-R[-1]C,N(A107030抵扣应纳税所得额明细表!R[-3]C[-1])),0)";
                                WorkingPaper.Wb.Worksheets["抵扣应纳税所得额审核表"].Range["C10"].Formula =
                                    @"= MAX('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!$D$22 - 'A100000中华人民共和国企业所得税年度纳税申报表（A类）'!$D$23 -$C$18, 0)";
                                WorkingPaper.Wb.Worksheets["抵扣应纳税所得额审核表"].Range["C18"].Formula =
                                    @"=MAX(0,MIN(C14,C17,'A100000中华人民共和国企业所得税年度纳税申报表（A类）'!$D$22-'A100000中华人民共和国企业所得税年度纳税申报表（A类）'!$D$23))";
                                #endregion

                                #region 表格名称
                                WorkingPaper.Wb.Worksheets["企业所得税年度纳税申报表填报表单"].Name = "（三）企业所得税年度纳税申报表填报表单";
                                WorkingPaper.Wb.Worksheets["（六）企业各税（费）审核汇总表"].Name = "（四）企业各税（费）审核汇总表";
                                WorkingPaper.Wb.Worksheets["（七）社会保险费明细表"].Name = "（五）社会保险费明细表";
                                #endregion

                                #region 移动表格
                                WorkingPaper.Wb.Worksheets["检查表"].Move(Before: WorkingPaper.Wb.Worksheets["余额表"]);
                                WorkingPaper.Wb.Worksheets[new string[] { "(三)子表12.企业所得税汇总纳税分支机构所得税分配表",
                "(四)无限期结转扣除项目情况表", "(五)跨年度确认所得情况表" }].Move(Before: WorkingPaper.Wb.Worksheets["基本情况（封面）"]);
                                WorkingPaper.Wb.Sheets.Add(After: WorkingPaper.Wb.Worksheets["A109010企业所得税汇总纳税分支机构所得税分配表"],
                                    Type: AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "\\0504新增表.xlsx");
                                WorkingPaper.Wb.Worksheets["A110010特殊性处理报告表"].Range["B3:C3"].Formula = "=基本情况!B2";
                                WorkingPaper.Wb.Worksheets["A110010特殊性处理报告表"].Range["B4:C4"].Formula = "=基本情况!B31";
                                WorkingPaper.Wb.Worksheets["A110010特殊性处理报告表"].Range["B5:C5"].Formula = "=地税、基本情况!B4";
                                WorkingPaper.Wb.Worksheets["A110010特殊性处理报告表"].Range["B15:C15"].Formula = "=基本情况!B12";
                                WorkingPaper.Wb.Worksheets["A110010特殊性处理报告表"].Range["E3:G3"].Formula = "=基本情况!B48";
                                WorkingPaper.Wb.Worksheets["A110010特殊性处理报告表"].Range["E4:G4"].Formula = "=基本情况!B36";
                                WorkingPaper.Wb.Worksheets["A110010特殊性处理报告表"].Range["E5:G5"].Formula = "=基本情况!B34";
                                WorkingPaper.Wb.Worksheets["A110010特殊性处理报告表"].Range["E15:G15"].Formula = "=基本情况!B21";
                                WorkingPaper.Wb.Worksheets["A110011债务重组报告表"].Range["B24:C24"].Formula = "=基本情况!B12";
                                WorkingPaper.Wb.Worksheets["A110011债务重组报告表"].Range["E24"].Formula = "=基本情况!B21";
                                WorkingPaper.Wb.Worksheets["A110012股权收购报告表 "].Range["B59:D59"].Formula = "=基本情况!B12";
                                WorkingPaper.Wb.Worksheets["A110012股权收购报告表 "].Range["F59:I59"].Formula = "=基本情况!B21";
                                WorkingPaper.Wb.Worksheets["A110013资产收购报告表"].Range["B25:E25"].Formula = "=基本情况!B12";
                                WorkingPaper.Wb.Worksheets["A110013资产收购报告表"].Range["G25:I25"].Formula = "=基本情况!B21";
                                WorkingPaper.Wb.Worksheets["A110014企业合并报告表 "].Range["B25"].Formula = "=基本情况!B12";
                                WorkingPaper.Wb.Worksheets["A110014企业合并报告表 "].Range["D25"].Formula = "=基本情况!B21";
                                WorkingPaper.Wb.Worksheets["A110015企业分立申报表"].Range["B27:C27"].Formula = "=基本情况!B12";
                                WorkingPaper.Wb.Worksheets["A110015企业分立申报表"].Range["E27:F27"].Formula = "=基本情况!B21";
                                WorkingPaper.Wb.Worksheets["A110016非货币资产投资递延纳税调整表"].Range["B10:C10"].Formula = "=基本情况!B12";
                                WorkingPaper.Wb.Worksheets["A110016非货币资产投资递延纳税调整表"].Range["E10:F10"].Formula = "=基本情况!B21";
                                WorkingPaper.Wb.Worksheets["A110017居民企业资产（股权）划转特殊性税务处理申报表"].Range["B29:C29"].Formula = "=基本情况!B12";
                                WorkingPaper.Wb.Worksheets["A110017居民企业资产（股权）划转特殊性税务处理申报表"].Range["E29:H29"].Formula = "=基本情况!B21";
                                WorkingPaper.Wb.Worksheets[new string[] { "分支机构企业所得税申报表（A类）",
                "（四）企业各税（费）审核汇总表", "（五）社会保险费明细表" }].Move(After: WorkingPaper.Wb.Worksheets["A110017居民企业资产（股权）划转特殊性税务处理申报表"]);
                                WorkingPaper.Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Rows["44:51"].Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                                string[,] XinBH = new string[,] { { "A110010" }, { "A110011" }, { "A110012" }, { "A110013" }, { "A110014" }, { "A110015" }, { "A110016" }, { "A110017" } };
                                string[,] XinMC = new string[,] { { "    特殊性处理报告表" }, { "    债务重组报告表" }, { "    股权收购报告表" }, { "    资产收购报告表" }, { "    企业合并报告表" }, { "    企业分立报告表" }, { "    非货币资产投资递延纳税调整表 " }, { "    居民企业资产（股权）划转特殊性税务处理申报表" } };
                                string[,] XinBM = new string[,] { { "#A110010特殊性处理报告表!A1" }, { "#A110011债务重组报告表!A1" }, { "#'A110012股权收购报告表 '!A1" }, { "#A110013资产收购报告表!A1" }, { "#'A110014企业合并报告表 '!A1" }, { "#A110015企业分立申报表!A1" }, { "#A110016非货币资产投资递延纳税调整表!A1" }, { "#'A110017居民企业资产（股权）划转特殊性税务处理申报表'!A1" } };
                                WorkingPaper.Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["A44:A51"].Value2 = XinBH;
                                WorkingPaper.Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["C44:C51"].FormulaR1C1 = "=RC[3]";
                                WorkingPaper.Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["F44:F51"].Value2 = "否";
                                for (int j = 0; j <= 7; j++)
                                {
                                    WorkingPaper.Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Hyperlinks.Add(
                                        WorkingPaper.Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B" + (44 + j).ToString()],
                                        XinBM[j, 0], Type.Missing, XinMC[j, 0], XinMC[j, 0]);
                                    WorkingPaper.Wb.Worksheets["主页"].Hyperlinks.Add(
                                        WorkingPaper.Wb.Worksheets["主页"].Range["G" + (21 + j).ToString()],
                                        XinBM[j, 0], Type.Missing, XinMC[j, 0], XinMC[j, 0]);
                                }

                                #endregion

                                #endregion
                                Banben = "V20160504-0316";
                            }
                            if(Banben.Substring(0,9)=="V20160504")
                            {
                                WorkingPaper.Wb.Worksheets["签发单"].Range["A1:E1"].Value2 = "中汇百邦（厦门）税务师事务所有限公司";
                                WorkingPaper.Wb.Worksheets["业务约定"].Range["A4:G4"].Value2 = "受托方：  中汇百邦（厦门）税务师事务所有限公司   （以下简称乙方）";
                                WorkingPaper.Wb.Worksheets["业务约定"].Range["A28:G28"].Value2 = "    甲方（签章）：                         乙方（签章）：中汇百邦（厦门）税务师事务所有限公司";
                                WorkingPaper.Wb.Worksheets["当局声明"].Range["A3"].Value2 = "中汇百邦（厦门）税务师事务所有限公司：";
                                WorkingPaper.Wb.Worksheets["报告封面"].Range["A1:G1"].Value2 = "BaiBang中汇百邦（厦门）税务师事务所有限公司";
                                WorkingPaper.Wb.Worksheets["报告封面"].Range["A27:G27"].Value2 = "                          审计单位： 中汇百邦（厦门）税务师事务所有限公司";
                                WorkingPaper.Wb.Worksheets["报告正文"].Range["A56:D56"].Value2 = "中汇百邦（厦门）税务师事务所有限公司";

                                WorkingPaper.Wb.Worksheets["A000000企业基础信息表"].Range["B15:F15"].Value2 = "月末一次加权平均法";
                                WorkingPaper.Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["F31"].Formula =
                                    @"=IF(SUM(A107014研发费用加计扣除优惠明细表!T16)<>0,""是"",""否"")";
                                WorkingPaper.Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["F18"].Formula =
                                    @"=IF(A105070捐赠支出纳税调整明细表!C25+A105070捐赠支出纳税调整明细表!G25<>0,""是"",""否"")";
                                WorkingPaper.Wb.Worksheets["A000000企业基础信息表"].Range["E20"].Value2 = "比例(%)";
                                WorkingPaper.Wb.Worksheets["A000000企业基础信息表"].Range["D27"].Value2 = "比例(%)";
                                WorkingPaper.Wb.Worksheets["A000000企业基础信息表"].Range["E21:E25"].FormulaR1C1 =
                                    @"=IFERROR(ROUND(基本情况!R[31]C[-2]*100,2),"""")";

                                WorkingPaper.Wb.Worksheets["A107011股息红利优惠明细表"].Range["E3:E4"].Value2 = "投资比例(%)";
                                WorkingPaper.Wb.Worksheets["A107011股息红利优惠明细表"].Range["L4"].Value2 = "减少投资比例(%)";
                                WorkingPaper.Wb.Worksheets["A107011股息红利优惠明细表"].Range["F6:F14"].FormulaR1C1 =
                                    @"=IF(RC[-4]="""","""",TEXT(股息红利优惠审核表!RC,""yyyy-mm-dd""))";

                                WorkingPaper.Wb.Worksheets["股息红利优惠审核表"].Range["E3:E4"].Value2 = "投资比例(%)";
                                WorkingPaper.Wb.Worksheets["股息红利优惠审核表"].Range["L4"].Value2 = "减少投资比例(%)";
                                WorkingPaper.Wb.Worksheets["股息红利优惠审核表"].Range["F6:F14"].NumberFormatLocal = "yyyy-mm-dd";
                                WorkingPaper.Wb.Worksheets["股息红利优惠审核表"].Range["F6:F14"].ShrinkToFit = true;

                                WorkingPaper.Wb.Worksheets["综合利用资源生产产品取得的收入优惠审核表"].Range["H3:H4"].Value2 = "综合利用的资源占生产产品材料的比例(%)";
                                WorkingPaper.Wb.Worksheets["综合利用资源生产产品取得的收入优惠审核表"].Range["C6:C14"].NumberFormatLocal = "yyyy-mm-dd";
                                WorkingPaper.Wb.Worksheets["综合利用资源生产产品取得的收入优惠审核表"].Range["C6:C14"].ShrinkToFit = true;

                                WorkingPaper.Wb.Worksheets["A107012综合利用资源生产产品取得的收入优惠明细表"].Range["B6:B14,D6:G14,I6:I14"].FormulaR1C1 =
                                    @"=综合利用资源生产产品取得的收入优惠审核表!RC&""""";
                                WorkingPaper.Wb.Worksheets["A107012综合利用资源生产产品取得的收入优惠明细表"].Range["C6:C1"].FormulaR1C1 =
                                    @"=IF(RC[-1]="""","""",TEXT(综合利用资源生产产品取得的收入优惠审核表!RC,""yyyy-mm-dd""))";
                                WorkingPaper.Wb.Worksheets["A107012综合利用资源生产产品取得的收入优惠明细表"].Range["H6:H14,J6:K14"].FormulaR1C1 =
                                    @"=IF(RC2="""","""",综合利用资源生产产品取得的收入优惠审核表!RC)";
                                WorkingPaper.Wb.Worksheets["A107012综合利用资源生产产品取得的收入优惠明细表"].Range["H3:H4"].Value2 = "综合利用的资源占生产产品材料的比例(%)";
                                WorkingPaper.Wb.Worksheets["高新技术企业优惠情况审核表"].Range["D14"].Formula =
                                    @"=IFERROR(ROUND(D10/D13*100,2),0)";
                                WorkingPaper.Wb.Worksheets["高新技术企业优惠情况审核表"].Range["C33"].Value2 = "十、本年研发费用占销售（营业）收入比例(%)";

                                WorkingPaper.Wb.Worksheets["软件、集成电路企业优惠情况审核表"].Range["E14"].Formula =
                                    @"=IFERROR(ROUND(E12/E11*100,2),0)";
                                WorkingPaper.Wb.Worksheets["软件、集成电路企业优惠情况审核表"].Range["E15"].Formula =
                                    @"=IFERROR(ROUND(E13/E11*100,2),0)";
                                WorkingPaper.Wb.Worksheets["软件、集成电路企业优惠情况审核表"].Range["E18"].Formula =
                                    @"=IFERROR(ROUND(E17/E16*100,2),0)";
                                WorkingPaper.Wb.Worksheets["软件、集成电路企业优惠情况审核表"].Range["E21"].Formula =
                                    @"=IFERROR(ROUND(E19/E16*100,2),0)";
                                WorkingPaper.Wb.Worksheets["软件、集成电路企业优惠情况审核表"].Range["E22"].Formula =
                                    @"=IFERROR(ROUND(E20/$E$16*100,2),0)";
                                WorkingPaper.Wb.Worksheets["软件、集成电路企业优惠情况审核表"].Range["E27"].Formula =
                                    @"=IFERROR(ROUND(E23/$E$16*100,2),0)";
                                WorkingPaper.Wb.Worksheets["软件、集成电路企业优惠情况审核表"].Range["E28"].Formula =
                                    @"=IFERROR(ROUND(E24/$E$16*100,2),0)";
                                WorkingPaper.Wb.Worksheets["软件、集成电路企业优惠情况审核表"].Range["E29"].Formula =
                                    @"=IFERROR(ROUND(E25/$E$16*100,2),0)";
                                WorkingPaper.Wb.Worksheets["软件、集成电路企业优惠情况审核表"].Range["E30"].Formula =
                                    @"=IFERROR(ROUND(E26/$E$16*100,2),0)";
                                WorkingPaper.Wb.Worksheets["软件、集成电路企业优惠情况审核表"].Range["E34"].Formula =
                                    @"=IFERROR(ROUND(E32/E31*100,2),0)";
                                WorkingPaper.Wb.Worksheets["软件、集成电路企业优惠情况审核表"].Range["E37"].Formula =
                                    @"=IFERROR(ROUND(E36/E35*100,2),0)";
                                WorkingPaper.Wb.Worksheets["软件、集成电路企业优惠情况审核表"].Range["E41"].Formula =
                                    @"=IFERROR(ROUND(E39/E38*100,2),0)";
                                WorkingPaper.Wb.Worksheets["软件、集成电路企业优惠情况审核表"].Range["E42"].Formula =
                                    @"=IFERROR(ROUND(E40/E39*100,2),0)";
                                WorkingPaper.Wb.Worksheets["软件、集成电路企业优惠情况审核表"].Range["E44"].Formula =
                                    @"=IFERROR(ROUND(E43/E39*100,2),0)";
                                WorkingPaper.Wb.Worksheets["软件、集成电路企业优惠情况审核表"].Range["D33"].Value2 =
                                    "十八、研究开发费用总额占企业销售（营业）收入总额的比例(%)";

                                WorkingPaper.Wb.Worksheets["A107014研发费用加计扣除优惠明细表"].Range["B6:B14"].FormulaR1C1 =
                                    @"=研发费用加计扣除优惠审核表!RC&""""";
                                WorkingPaper.Wb.Worksheets["A107014研发费用加计扣除优惠明细表"].Range["C6:T14"].FormulaR1C1 =
                                    @"=IF(RC2="""","""",研发费用加计扣除优惠审核表!RC)";
                                WorkingPaper.Wb.Worksheets["A107014研发费用加计扣除优惠明细表"].Range["C15:T15"].FormulaR1C1 =
                                    @"=研发费用加计扣除优惠审核表!RC";
                                WorkingPaper.Wb.Worksheets["A107041高新技术企业优惠情况及明细表"].Range["D5"].Formula =
                                    @"=IF(高新技术企业优惠情况审核表!D5="""","""",TEXT(高新技术企业优惠情况审核表!D5,""yyyy-mm-dd""))";
                                WorkingPaper.Wb.Worksheets["A107041高新技术企业优惠情况及明细表"].Range["D28"].Formula =
                                    @"=高新技术企业优惠情况审核表!D28";

                                WorkingPaper.Wb.Worksheets["主页"].Range["H3"].Formula = "=首页!A1";
                                WorkingPaper.Wb.Worksheets["主页"].Hyperlinks.Add(
                                    WorkingPaper.Wb.Worksheets["主页"].Range["F7"],
                                    "#'（三）企业所得税年度纳税申报表填报表单'!A1", Type.Missing, "（三）企业所得税年度纳税申报表填报表单",
                                    "（三）企业所得税年度纳税申报表填报表单");
                                WorkingPaper.Wb.Worksheets["主页"].Hyperlinks.Add(
                                    WorkingPaper.Wb.Worksheets["主页"].Range["J24"],
                                    "#'（四）企业各税（费）审核汇总表'!A1", Type.Missing, "（四）企业各税（费）审核汇总表",
                                    "（四）企业各税（费）审核汇总表");
                                WorkingPaper.Wb.Worksheets["主页"].Hyperlinks.Add(
                                    WorkingPaper.Wb.Worksheets["主页"].Range["J25"],
                                    "#'（五）社会保险费明细表'!A1", Type.Missing, "（五）社会保险费明细表",
                                    "（五）社会保险费明细表");
                                WorkingPaper.Wb.Worksheets["利润"].Range["C27"].Formula = "=C5-C6-C7-C15-C18-C22+C26-C24+C25";
                                WorkingPaper.Wb.Worksheets["利润"].Range["D27"].Formula = "=D5-D6-D7-D15-D18-D22+D26-D24+D25";
                                //WorkingPaper.Wb.Worksheets["A000000企业基础信息表"].Range["B6"].Formula = @"=ROUND(SUBSTITUTE(地税、基本情况!B9,""" + Convert.ToString(63) + @","""")/10000,2)";
                                WorkingPaper.Wb.Worksheets["A000000企业基础信息表"].Range["B8"].Formula =
                                "=ROUND(IF(截止月=起始月,0,SUM(OFFSET(社保明细工资人数!J8,VALUE(起始月),0,截止月-起始月,1))+OFFSET(社保明细工资人数!J7,VALUE(起始月),0)/2+OFFSET(社保明细工资人数!J8,VALUE(截止月),0)/2)/(截止月-起始月+1),0)";
                                WorkingPaper.Wb.Worksheets["A000000企业基础信息表"].Range["B9"].Formula =
                                    @"=ROUND(IF(AND(起始月=""01"",截止月=""12""),(资产负债!C35/2+资产负债!I20+资产负债!D35/2)/IF(资产负债!I20=0,1,12),IF(AND(起始月<>""01"",截止月<>""12""),(资产负债!C35/2+资产负债!I20+资产负债!D35/2)/IF(资产负债!I20=0,1/2,截止月-起始月+1.5),IF(截止月<>""12"",(资产负债!C35/2+资产负债!I20+资产负债!D35/2)/IF(资产负债!I20=0,1,截止月+1),IF(起始月<>""01"",(资产负债!C35/2+资产负债!I20+资产负债!D35/2)/IF(资产负债!I20=0,1/2,截止月-起始月+0.5)))))/10000,2)";
                                WorkingPaper.Wb.Worksheets["A000000企业基础信息表"].Range["B6"].Formula =
                                @"=IFERROR(ROUND(SUBSTITUTE(地税、基本情况!B9,税金申报明细!R5,"""")/10000,2),ROUND(地税、基本情况!B9/10000,2))";
                                if (Banben == "V20160504-0316")
                                {
                                    //WorkingPaper.Wb.Worksheets["A107014研发费用加计扣除优惠明细表"].Rows["15:15"].Delete(XlDeleteShiftDirection.xlShiftUp);
                                    //WorkingPaper.Wb.Worksheets["研发费用加计扣除优惠审核表"].Rows["15:15"].Delete(XlDeleteShiftDirection.xlShiftUp);
                                    WorkingPaper.Wb.Worksheets["A107040减免所得税优惠明细表"].Columns["C:C"].Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                                    WorkingPaper.Wb.Worksheets["资产负债"].Range["I8"].Value2 = "资产总额";
                                    WorkingPaper.Wb.Worksheets["资产负债"].Range["J8:J19"].Value2 =
                                        new string[,] { { "月份" }, { "1月底" }, { "2月底" }, { "3月底" }, { "4月底" }, { "5月底" }, { "6月底" }, { "7月底" }, { "8月底" }, { "9月底" }, { "10月底" }, { "11月底" } };
                                    WorkingPaper.Wb.Worksheets["资产负债"].Range["I20"].Formula = "=SUM(I9:I19)";
                                    WorkingPaper.Wb.Worksheets["资产负债"].Range["I8,J8:J19"].Interior.Color = 13434828;//绿色
                                    WorkingPaper.Wb.Worksheets["资产负债"].Range["I20"].Interior.Color = 12632256;//灰色
                                    Banben = "V20160508-0504-0316";
                                }
                                else
                                    Banben = "V20160508-0504";
                            }
                            WorkingPaper.Wb.Worksheets["首页"].Range["A1"].Value2 = Banben;
                            Globals.WPToolAddln.Application.StatusBar = false;
                            MessageBox.Show("升级完成，请检查！");
                        }
                    }
                }
                else
                {
                    MessageBox.Show("当前版本为："+Banben+"，最新版本为：V20160504。不需要升级", "提示！",
                        MessageBoxButtons.OK);
                }
            }

        }

        private void btn工具设置_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnOUT07_Click(object sender, RibbonControlEventArgs e)
        {
            if (WorkingPaper.OOO)
            {
                if (MessageBox.Show("现在将当前可见工作表导出为07版本Excel。是否继续？", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    SaveFileDialog Sv = new SaveFileDialog();
                    Sv.Filter = "Excel 2007工作簿(*.xlsx)|*.xlsx";
                    Sv.FileName = "税审工作表导出";
                    Sv.Title = "导出当前可见工作表";
                    Sv.OverwritePrompt = true;
                    Sv.InitialDirectory = WorkingPaper.Wb.Path;
                    //Sv.RestoreDirectory = true;
                    if (Sv.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            Globals.WPToolAddln.Application.StatusBar = "正在导出可见工作表...";
                            Globals.WPToolAddln.Application.DisplayAlerts = false;
                            Globals.WPToolAddln.Application.ScreenUpdating = false;
                            Workbook Oldbook = WorkingPaper.Wb;
                            List<string> lists = new List<string>();
                            foreach (Worksheet ssh in Oldbook.Worksheets)
                            {
                                if (ssh.Visible == XlSheetVisibility.xlSheetVisible)
                                {
                                    lists.Add(ssh.Name.ToString());
                                }
                            }
                            string[] s = lists.ToArray();
                            Workbook Newbook = Globals.WPToolAddln.Application.Workbooks.Add();
                            int C = Newbook.Worksheets.Count;
                            Oldbook.Worksheets[s].Copy(Type.Missing, Newbook.Worksheets[C]);
                            for (int i = 1; i <= C; i++)
                            {
                                Newbook.Worksheets[1].Delete();
                            }
                            foreach (Name nm in Newbook.Names)
                            {
                                if (Regex.IsMatch(nm.RefersTo.ToString(), @"(#REF!)|\/|\\|\*|\[|\]"))
                                {
                                    nm.Delete();
                                }
                            }
                            Newbook.BreakLink(Oldbook.Path + "\\" + Oldbook.Name, XlLinkType.xlLinkTypeExcelLinks);
                            Newbook.SaveAs(Sv.FileName.ToString(), XlFileFormat.xlOpenXMLWorkbook);
                            Newbook.Close();
                            Newbook = null;
                            Globals.WPToolAddln.Application.DisplayAlerts = true;
                            Globals.WPToolAddln.Application.ScreenUpdating = true;
                            Globals.WPToolAddln.Application.StatusBar = false;
                            MessageBox.Show("文件导出完成！");
                        }
                        catch (Exception ex)
                        {
                            Globals.WPToolAddln.Application.DisplayAlerts = true;
                            Globals.WPToolAddln.Application.ScreenUpdating = true;
                            Globals.WPToolAddln.Application.StatusBar = false;
                            MessageBox.Show("用户操作出现错误：" + ex.Message);
                        }
                    }

                }
            }
        }

        private void btn导出报告_Click(object sender, RibbonControlEventArgs e)
        {
            if (WorkingPaper.OOO)
            {
                if (Math.Round(CU.Shuzi(WorkingPaper.Wb.Worksheets["A107040减免所得税优惠明细表"].Range["D7"].Value2) +
                    CU.Shuzi(WorkingPaper.Wb.Worksheets["A107040减免所得税优惠明细表"].Range["D8"].Value2), 2) !=
                    Math.Round(CU.Shuzi(WorkingPaper.Wb.Worksheets["A107040减免所得税优惠明细表"].Range["D6"].Value2),2))
                {
                    MessageBox.Show("A107040减免所得税优惠明细表，D6不等于D7+D8，请检查后重试。");
                    return;
                }
                    
                if (MessageBox.Show("现在要导出上传报告文件。是否继续？", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    SaveFileDialog Sv = new SaveFileDialog();
                    Sv.Filter = "Excel 2003工作簿(*.xls)|*.xls";
                    Sv.FileName = "上传报告导出";
                    Sv.Title = "导出上传报告";
                    Sv.OverwritePrompt = true;
                    Sv.InitialDirectory = WorkingPaper.Wb.Path;
                    //Sv.RestoreDirectory = true;
                    if (Sv.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            Globals.WPToolAddln.Application.StatusBar = "正在导出报告...";
                            Globals.WPToolAddln.Application.DisplayAlerts = false;
                            Globals.WPToolAddln.Application.ScreenUpdating = false;
                            File.Copy(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "\\上传报告.xls", Sv.FileName.ToString(), true);
                            CU.事项说明();
                            Workbook Newbook = Globals.WPToolAddln.Application.Workbooks.Open(Sv.FileName.ToString(), XlUpdateLinks.xlUpdateLinksNever);
                            Newbook.ChangeLink(Name: @"E:\税审底稿 模板.xlsx", NewName: WorkingPaper.Wb.FullName, Type: XlLinkType.xlLinkTypeExcelLinks);
                            //Newbook.UpdateLink(WorkingPaper.Wb.FullName, XlLinkType.xlLinkTypeExcelLinks);
                            Newbook.BreakLink(WorkingPaper.Wb.FullName, XlLinkType.xlLinkTypeExcelLinks);
                            Worksheet SH = WorkingPaper.Wb.Sheets["(二)附表-纳税调整额的审核"];
                            object[,] Arr = SH.Range["A7:E" + SH.Cells[SH.UsedRange.Rows.Count + 1, 1].End[XlDirection.xlUp].Row.ToString()].Value2;
                            Newbook.Worksheets["(二)附表-纳税调整额的审核"].Range["A7:E" + SH.Cells[SH.UsedRange.Rows.Count + 1, 1].End[XlDirection.xlUp].Row.ToString()].Value2 = Arr;
                            Newbook.Save();
                            Newbook.Close();
                            Newbook = null;
                            Globals.WPToolAddln.Application.DisplayAlerts = true;
                            Globals.WPToolAddln.Application.ScreenUpdating = true;
                            Globals.WPToolAddln.Application.StatusBar = false;
                            MessageBox.Show("上传报告导出完成！");
                        }
                        catch (Exception ex)
                        {
                            Globals.WPToolAddln.Application.DisplayAlerts = true;
                            Globals.WPToolAddln.Application.ScreenUpdating = true;
                            Globals.WPToolAddln.Application.StatusBar = false;
                            MessageBox.Show("用户操作出现错误：" + ex.Message);
                        }
                    }

                }
            }
        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            if(WorkingPaper.OOO)
            {
                if (MessageBox.Show("是否自动修复 A107014研发费用加计扣除优惠明细表 和 研发费用加计扣除优惠审核表 错误？", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    bool ll = false;
                    if (CU.Zifu(WorkingPaper.Wb.Worksheets["研发费用加计扣除优惠审核表"].Range["A15"].Value2) != "10（1+2+3+4+5+6+7+8+9）")
                    {
                        Worksheet SH = WorkingPaper.Wb.Worksheets["研发费用加计扣除优惠审核表"];
                        SH.Range["A15:B15,U15"].Interior.Color = 13434828;//绿色
                        SH.Range["C15:T15"].Interior.Color = 12632256;//灰色
                        SH.Range["A15"].Value2 = "10（1+2+3+4+5+6+7+8+9）";
                        SH.Range["B15"].Value2 = "合计";
                        SH.Range["U15"].Value2 = "第10行第19列＝表A107010第22行";
                        SH.Range["C15:T15"].FormulaR1C1 = "=SUM(R[-9]C:R[-1]C)";
                        SH.Range["A15:U15"].Font.Size = 10;
                        SH.Range["A15:U15"].Borders.LineStyle = XlLineStyle.xlContinuous;
                        SH.Range["A15,U15"].WrapText = true;
                        ll = true;
                    }
                    if (CU.Zifu(WorkingPaper.Wb.Worksheets["A107014研发费用加计扣除优惠明细表"].Range["A15"].Value2) != "10（1+2+3+4+5+6+7+8+9）")
                    {
                        Worksheet SH = WorkingPaper.Wb.Worksheets["A107014研发费用加计扣除优惠明细表"];
                        SH.Range["A15:B15,U15"].Interior.Color = 13434828;//绿色
                        SH.Range["C15:T15"].Interior.Color = 16764057;//蓝色
                        SH.Range["A15"].Value2 = "10（1+2+3+4+5+6+7+8+9）";
                        SH.Range["B15"].Value2 = "合计";
                        SH.Range["U15"].Value2 = "第10行第19列＝表A107010第22行";
                        SH.Range["C15:T15"].FormulaR1C1 = "=研发费用加计扣除优惠审核表!RC";
                        WorkingPaper.Wb.Worksheets["免税、减计收入及加计扣除优惠审核表"].Range["C25"].Formula = "=A107014研发费用加计扣除优惠明细表!T15";
                        SH.Range["A15:U15"].Font.Size = 10;
                        SH.Range["A15:U15"].Borders.LineStyle = XlLineStyle.xlContinuous;
                        SH.Range["A15,U15"].WrapText = true;
                        ll = true;
                    }
                    if (ll)
                        MessageBox.Show("修复完成，请检查。");
                    else
                        MessageBox.Show("此底稿无需修复。");
                }
            }
        }

        private void btn底稿打印_Click(object sender, RibbonControlEventArgs e)
        {
            底稿打印 dgdy = new 底稿打印();
            dgdy.ShowDialog();
    //        Range("A1:K27").Select
    //Selection.Copy
    //Sheets("Sheet2").Select
    //ActiveSheet.Shapes.AddShape(, 51.6, 25.8, 72#, 72#).Select
        }
    }
}
