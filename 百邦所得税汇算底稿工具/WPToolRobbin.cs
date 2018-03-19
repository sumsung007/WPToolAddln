using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Core;
using System.Text.RegularExpressions;

namespace 百邦所得税汇算底稿工具
{
    public partial class WorkingPaper
    {
        Microsoft.Office.Tools.CustomTaskPane Excel10Taskpane;
        public static Workbook Wb,wb打印;
        public static Boolean OOO=false;
        public static int 版本号;
        public static string 当前版本 = "20171222修复版"; //Assembly.GetExecutingAssembly().GetName().Version.ToString().Replace(".", "");
        public static int Excel版本;

        public Dictionary<int, Microsoft.Office.Tools.CustomTaskPane> TaskPanels =
            new Dictionary<int, Microsoft.Office.Tools.CustomTaskPane>();

        public Dictionary<int, Contents> Cons =
            new Dictionary<int, Contents>();
        Contents Excel10Con;
        CommandBarButton Cd;
        //

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Contact.Label = "联系我们\n";
            splitButton1.Label = "报告导出\n";
            sb导出数据.Label = "导出数据\n";
            btn工具设置.Label = "高级功能\n";
            btn检查表.Label = "检查表\n";
            switch (Globals.WPToolAddln.Application.Version)
            {
                case "15.0":
                case "16.0":
                    Excel版本 = 13;
                    break;
                case "14.0":
                    Excel版本 = 10;
                    break;
                case "12.0":
                    Excel版本 = 07;
                    break;
            }
            if (Excel版本 == 10|| Excel版本 == 07)
            {
                Excel10Con = new Contents();
                Excel10Taskpane = Globals.WPToolAddln.CustomTaskPanes.Add(Excel10Con, "税审底稿工具");
                Excel10Taskpane.Width = 300;
                Excel10Taskpane.VisibleChanged += new EventHandler(MyTaskpane_VisibleChanged);
            }
            if (!CU.授权检测())
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
                splitButton1.Enabled = false;
                btn注册.Visible = true;
                MessageBox.Show("底稿工具尚未注册，请进入设置后将机器码发给授权单位授权！");
            }
            else
            {
                if (Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\BaiBang", "NewVersion", String.Empty)
                        .ToString() !=
                    当前版本)
                {
                    AboutBox1 ab = new AboutBox1();
                    ab.ShowDialog();
                    Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\BaiBang", "NewVersion", 当前版本);
                }
                if (Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\BaiBang", "Updatatime", String.Empty)
        .ToString() !=DateTime.Now.Date.ToShortDateString())
                {
                    更新(false);
                    Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\BaiBang", "Updatatime", DateTime.Now.Date.ToShortDateString());
                }
            }

            Globals.WPToolAddln.Application.WorkbookActivate += Application_WorkbookActivate;
            
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
                    if (WorkingPaper.版本号 == 2018)
                    {
                        WorkingPaper.Wb.Application.ScreenUpdating = false;
                        WorkingPaper.Wb.Worksheets[1].Visible = XlSheetVisibility.xlSheetVisible;
                        int C = WorkingPaper.Wb.Worksheets.Count;
                        for (int i = 2; i <= C; i++)
                        {
                            WorkingPaper.Wb.Worksheets[i].Visible = XlSheetVisibility.xlSheetVisible;
                        }
                        WorkingPaper.Wb.Worksheets[1].Visible = XlSheetVisibility.xlSheetHidden;
                        WorkingPaper.Wb.Sheets["开始"].Select();
                        WorkingPaper.Wb.Application.ScreenUpdating = true;

                    }
                    else
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
                }
                catch (Exception ex)
                {
                    Globals.WPToolAddln.Application.ScreenUpdating = true;
                    MessageBox.Show("用户操作出现错误：" + ex.Message);
                }
            }
        }

        private void MyTaskpane_VisibleChanged(object sender, EventArgs e)
        {
            if (Excel版本 == 13)
            {
                int hwnd = Globals.WPToolAddln.Application.ActiveWindow.Hwnd;
                Microsoft.Office.Tools.CustomTaskPane mypane;
                TaskPanels.TryGetValue(hwnd, out mypane);
                if (mypane != null) tb显示目录.Checked = mypane.Visible;
            }
            else if (Excel10Taskpane != null) tb显示目录.Checked = Excel10Taskpane.Visible;

        }

        private void btnHelp_Click(object sender, RibbonControlEventArgs e)         //关于程序
        {
            AboutBox1 AB = new AboutBox1();
            AB.ShowDialog();
            
        }

        private void btn新建_Click(object sender, RibbonControlEventArgs e)          //新建底稿文件
        {
            
            DialogResult dr = MessageBox.Show("是否新建一个2017年版申报表税审底稿？按否新建2014年版底稿，按取消则不新建。", "新建", MessageBoxButtons.YesNoCancel);
            if (dr==DialogResult.Cancel)
            {
                return;
            }
            SaveFileDialog Sv = new SaveFileDialog();
            string sname = "\\税审底稿模板.xlsx";
            Sv.FileName = "税审2016年底稿";
            if (dr == DialogResult.Yes)
            {
                Sv.FileName = "税审2017年底稿";
                sname = "\\税审底稿2017模板.xlsx";

            }

            Sv.Filter = "税审底稿(*.xlsx)|*.xlsx";
                Sv.Title = "保存新的税审底稿";
                Sv.OverwritePrompt = true;
                Sv.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
            if (Sv.ShowDialog() == DialogResult.OK)
            {
                File.Copy(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + sname, Sv.FileName.ToString(),
                    true);
                Globals.WPToolAddln.Application.Workbooks.Open(Sv.FileName.ToString());
            }

        }
        

        private void tb显示目录_Click(object sender, RibbonControlEventArgs e)
        {
            if (Excel版本 == 13)
            {
                int hwnd = Globals.WPToolAddln.Application.ActiveWindow.Hwnd;
                Microsoft.Office.Tools.CustomTaskPane mypane;
                TaskPanels.TryGetValue(hwnd, out mypane);
                if (mypane != null)
                {
                    mypane.Visible = tb显示目录.Checked;
                }
                else
                {
                    Contents con = new Contents();
                    Microsoft.Office.Tools.CustomTaskPane pane = Globals.WPToolAddln.CustomTaskPanes.Add(con, "税审底稿工具",
                        Globals.WPToolAddln.Application.ActiveWindow);
                    //这一步很重要将决定是否显示到当前窗口，第三个参数的意思就是依附到那个窗口
                    //pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                    pane.Width = 300;
                    TaskPanels.Add(hwnd, pane);
                    pane.VisibleChanged += new EventHandler(MyTaskpane_VisibleChanged);
                    pane.Visible = tb显示目录.Checked;
                }
            }
            else
            {
                if (Excel10Taskpane == null)
                {
                    Excel10Con = new Contents();
                    Excel10Taskpane = Globals.WPToolAddln.CustomTaskPanes.Add(Excel10Con, "税审底稿工具");
                    Excel10Taskpane.Width = 300;
                    Excel10Taskpane.VisibleChanged += new EventHandler(MyTaskpane_VisibleChanged);
                }
                Excel10Taskpane.Visible = tb显示目录.Checked;
            }
        }

        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            if (WorkingPaper.OOO)
            {
                if (WorkingPaper.版本号 == 2018)
                {
                    CU.工作表切换(new string[] { "A100000 中华人民共和国企业所得税年度纳税申报表（A类）" ,
                        "A000000 企业基础信息表","A106000 企业所得税弥补亏损明细表" ,"调整事项","凭证检查",
                        "企业基本情况","交换意见","当局声明" ,"业务约定","现金证明"});
                    WorkingPaper.Wb.Sheets["企业基本情况"].Range["$H$21:$H$128"].AutoFilter(Field: 1, Criteria1: "=1"); 
                }
                else
                {

                    CU.工作表切换(new string[] { "A100000中华人民共和国企业所得税年度纳税申报表（A类）" ,
                        "A000000企业基础信息表","A106000企业所得税弥补亏损明细表" ,"事项说明","凭证检查",
                        "(二)附表-纳税调整额的审核","交换意见","当局声明" ,"业务约定","现金证明"});
                    CU.事项说明();
                }

            }
        }



        //菜单按键
        private void btn基本情况_Click(object sender, RibbonControlEventArgs e)
        {
            if (OOO)
            {

                if (WorkingPaper.版本号 == 2018)
                {
                    CU.工作表切换(new string[] {"基本情况", "地税、基本情况", "A000000 企业基础信息表"});
                    Wb.Worksheets["基本情况"].Select();
                }
                else
                {

                    CU.工作表切换(new string[] {"基本情况", "地税、基本情况", "A000000企业基础信息表"});
                    Wb.Worksheets["基本情况"].Select();
                }
            }
        }

        private void btn余额报表_Click(object sender, RibbonControlEventArgs e)
        {
            if (WorkingPaper.版本号 == 2018)
            {
                CU.工作表切换(new string[] { "余额表", "资产负债表", "利润表" });
            }
            else
            {
                CU.工作表切换(new string[] { "余额表", "资产负债", "利润" });
            }
        }

        private void btn税费测算_Click(object sender, RibbonControlEventArgs e)
        {
            if (WorkingPaper.版本号 == 2018)
            {
                CU.工作表切换(new string[] { "应交税费","收入与申报核对表","税金附加","税费缴纳测算",
                    "社保明细工资人数","补亏","企业各税审核汇总表","税金申报明细"});
            }
            else
            {

                CU.工作表切换(new string[] { "纳税申报数据","主营税金","税费缴纳测算","纳税申报数据",
                    "收入与申报核对表","企业各税审核汇总表","税金申报明细","社保明细工资人数","补亏"});
            }

        }

        private void btn检查表_Click(object sender, RibbonControlEventArgs e)
        {
            if (WorkingPaper.版本号 == 2018)
            {

                CU.工作表切换(new string[] { "凭证检查", "检查表" });
                Wb.Sheets["检查表"].Rows["2:66"].Hidden = false;
                string s = "";
                double k;
                object[,] JCB = Wb.Sheets["检查表"].Range["C2:C66"].Value2;
                for (int i = 1; i <= 65; i++)
                {
                    if (JCB[i, 1] != null)
                    {
                        if (double.TryParse(JCB[i, 1].ToString().Trim(), out k))
                        {

                            if (k == 0)
                            {
                                s = s + ",C" + (i + 1).ToString();
                            }
                        }
                    }
                }
                if (s.Length > 0)
                {
                    Wb.Sheets["检查表"].Range[s.Substring(1, s.Length - 1)].EntireRow.Hidden = true;
                }
            }
            else
            {

                CU.工作表切换(new string[] { "凭证检查", "检查表" });
                Wb.Sheets["检查表"].Rows["2:73"].Hidden = false;
                string s = "";
                double k;
                object[,] JCB = Wb.Sheets["检查表"].Range["C2:C73"].Value2;
                for (int i = 1; i <= 72; i++)
                {
                    if (JCB[i, 1] != null)
                    {
                        if (double.TryParse(JCB[i, 1].ToString().Trim(), out k))
                        {

                            if (k == 0)
                            {
                                s = s + ",C" + (i + 1).ToString();
                            }
                        }
                    }
                }
                if (s.Length > 0)
                {
                    Wb.Sheets["检查表"].Range[s.Substring(1, s.Length - 1)].EntireRow.Hidden = true;
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
                splitButton1.Enabled = true;
                btn注册.Visible = false;
            }
        }

        private void btn底稿查看_Click(object sender, RibbonControlEventArgs e)
        {
            if (WorkingPaper.OOO)
            {
                try
                {

                    if (WorkingPaper.版本号 == 2018)
                    {
                        WorkingPaper.Wb.Application.ScreenUpdating = false;
                        WorkingPaper.Wb.Worksheets[1].Visible = XlSheetVisibility.xlSheetVisible;
                        int C = WorkingPaper.Wb.Worksheets.Count;
                        for (int i = 2; i <= C; i++)
                        {
                            WorkingPaper.Wb.Worksheets[i].Visible = XlSheetVisibility.xlSheetVisible;
                        }
                        WorkingPaper.Wb.Worksheets[1].Visible = XlSheetVisibility.xlSheetHidden;
                        WorkingPaper.Wb.Sheets["开始"].Select();
                        WorkingPaper.Wb.Application.ScreenUpdating = true;
                    }
                    else
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
            if (WorkingPaper.版本号 == 2018)
            {
                CU.工作表切换(new string[]
                {
                    "报告正文", "企业基本情况", "封面", "企业所得税年度纳税申报表填报表单", "A000000 企业基础信息表", "A100000 中华人民共和国企业所得税年度纳税申报表（A类）",
                    "A101010 一般企业收入明细表", "A101020 金融企业收入明细表", "A102010 一般企业成本支出明细表", "A102020 金融企业支出明细表",
                    "A103000事业单位、民间非营利组织收入、支出明细表", "A104000期间费用明细表", "A105000纳税调整项目明细表",
                    "A105010视同销售和房地产开发企业特定业务纳税调整明细表", "A105020未按权责发生制确认收入纳税调整明细表", "A105030投资收益纳税调整明细表",
                    "A105040专项用途财政性资金纳税调整表", "A105050职工薪酬支出及纳税调整明细表", "A105060广告费和业务宣传费跨年度纳税调整明细表",
                    "A105070捐赠支出及纳税调整明细表", "A105080 资产折旧、摊销及纳税调整明细表", "A105090资产损失税前扣除及纳税调整明细表",
                    "A105100企业重组及递延纳税事项调整明细表", "A105110政策性搬迁纳税调整明细表", "A105120 特殊行业准备金及纳税调整明细表", "A106000 企业所得税弥补亏损明细表",
                    "A107010免税、减计收入及加计扣除优惠明细表", "A107011符合条件的居民企业之间的股息、红利等…优惠明细表", "A107012 研发费用加计扣除优惠明细表",
                    "A107020所得减免优惠明细表", "A107030 抵扣应纳税所得额明细表", "A107040减免所得税优惠明细表", "A107041 高新技术企业优惠情况及明细表",
                    "A107042软件、集成电路企业优惠情况及明细表", "A107050 税额抵免优惠明细表", "A108000境外所得税收抵免明细表", "A108010境外所得纳税调整后所得明细表",
                    "A108020境外分支机构弥补亏损明细表", "A108030 跨年度结转抵免境外所得税明细表", "A109000跨地区经营汇总纳税企业年度分摊企业所得税明细表",
                    "A109010 企业所得税汇总纳税分支机构所得税分配表", "企业各税审核汇总表",
                });
                WorkingPaper.Wb.Sheets["报告正文"].Select();
                CU.事项说明();
            }
            else
            {

                CU.工作表切换(new string[]
                {
                    "报告封面", "报告正文", "基本情况（封面）", "1.保留意见", "2.否定意见", "3.无保留意见", "4.无法表明意见", "(二)企业基本情况和审核事项说明", "(二)附表-科目说明",
                    "(二)附表-纳税调整额的审核", "（三）企业所得税年度纳税申报表填报表单", "A000000企业基础信息表", "A100000中华人民共和国企业所得税年度纳税申报表（A类）",
                    "A101010一般企业收入明细表",
                    "A101020金融企业收入明细表", "A102010一般企业成本支出明细表", "A102020金融企业支出明细表", "A103000事业单位、民间非营利组织收入、支出明细表",
                    "A104000期间费用明细表",
                    "A105000纳税调整项目明细表", "A105010视同销售和房地产开发企业特定业务纳税调整明细表", "A105020未按权责发生制确认收入纳税调整明细表", "A105030投资收益纳税调整明细表",
                    "A105040专项用途财政性资金纳税调整表", "A105050职工薪酬纳税调整明细表", "A105060广告费和业务宣传费跨年度纳税调整明细表", "A105070捐赠支出纳税调整明细表",
                    "A105080资产折旧、摊销情况及纳税调整明细表",
                    "A105081固定资产加速折旧、扣除明细表", "A105090资产损失税前扣除及纳税调整明细表", "A105091资产损失（专项申报）税前扣除及纳税调整明细表",
                    "A105100企业重组纳税调整明细表",
                    "A105110政策性搬迁纳税调整明细表", "A105120特殊行业准备金纳税调整明细表", "A106000企业所得税弥补亏损明细表", "A107010免税、减计收入及加计扣除优惠明细表",
                    "A107011股息红利优惠明细表",
                    "A107012综合利用资源生产产品取得的收入优惠明细表", "A107013金融保险等机构取得涉农利息保费收入优惠明细表", "A107014研发费用加计扣除优惠明细表",
                    "A107020所得减免优惠明细表",
                    "A107030抵扣应纳税所得额明细表", "A107040减免所得税优惠明细表", "A107041高新技术企业优惠情况及明细表", "A107042软件、集成电路企业优惠情况及明细表",
                    "A107050税额抵免优惠明细表",
                    "A108000境外所得税收抵免明细表", "A108010境外所得纳税调整后所得明细表", "A108020境外分支机构弥补亏损明细表", "A108030跨年度结转抵免境外所得税明细表",
                    "A109000跨地区经营汇总纳税企业年度分摊企业所得税明细表",
                    "A109010企业所得税汇总纳税分支机构所得税分配表", "A110010特殊性处理报告表", "A110011债务重组报告表", "A110012股权收购报告表 ", "A110013资产收购报告表",
                    "A110014企业合并报告表 ", "A110015企业分立申报表", "A110016非货币资产投资递延纳税调整表", "A110017居民企业资产（股权）划转特殊性税务处理申报表",
                    "研发项目可加计扣除研究开发费用情况归集表", "（四）企业各税（费）审核汇总表", "（五）社会保险费明细表"
                });
                WorkingPaper.Wb.Sheets["基本情况（封面）"].Select();
                CU.事项说明();
            }

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
            MessageBox.Show(Wb.Name.Substring(0,Wb.Name.LastIndexOf(".")));
            /*
            if (WorkingPaper.OOO)
            {
                CU.工作表切换(new string[] { "A100000中华人民共和国企业所得税年度纳税申报表（A类）" ,
                "A000000企业基础信息表","A106000企业所得税弥补亏损明细表" ,"事项说明","凭证检查",
                "(二)附表-纳税调整额的审核","交换意见","当局声明" ,"业务约定"});
                CU.事项说明();
            }*/
        }

        private void 底稿升级_Click(object sender, RibbonControlEventArgs e)
        {
            if (WorkingPaper.OOO)
            {
                if (WorkingPaper.版本号 == 2018)
                {
                    MessageBox.Show("暂不需要升级。");
                }
                else
                {

                    string Banben1 = CU.Zifu(WorkingPaper.Wb.Worksheets["首页"].Range["A1"].Value2);
                    string Banben = "";
                    bool 升级 = false;
                    Banben = Banben1;
                    switch (Banben1.Substring(0, 9))
                    {
                        case "V20171222":
                            升级 = false;
                            break;
                        default:
                            升级 = true;
                            break;
                    }

                    if (升级)
                    {
                        if (MessageBox.Show("当前版本为：" + Banben + "，最新版本为：V20171222。是否升级？", "提示！",
                            MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            if (MessageBox.Show("本操作具有不稳定性，会先保存当前文件，并以BAK后缀文件备份在文件同目录下。是否继续？", "警告！",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                            {
                                Globals.WPToolAddln.Application.StatusBar = "正在升级底稿...";

                                WorkingPaper.Wb.Worksheets["首页"].Unprotect();
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

                                if (Banben.Substring(0, 9) == "V20170210")
                                {
                                    #region 20170210升级为20170312


                                    WorkingPaper.Wb.Worksheets["A000000企业基础信息表"].Range["B7"].NumberFormatLocal = "G/通用格式";
                                    WorkingPaper.Wb.Worksheets["A000000企业基础信息表"].Range["B7"].Formula = "=LEFT(地税、基本情况!F6,4)";
                                    WorkingPaper.Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["F17"].Formula =
                                        "=IF(OR(A105060广告费和业务宣传费跨年度纳税调整明细表!C4<>0,A105060广告费和业务宣传费跨年度纳税调整明细表!C11<>0,A105060广告费和业务宣传费跨年度纳税调整明细表!C15<>0),\"是\",\"否\")";

                                    //福利费和业务招待费调整
                                    WorkingPaper.Wb.Worksheets["制造费用、生产成本"].Range["F23"].Formula =
                                        "=-SUMIFS(凭证检查!G6:G205,凭证检查!E6:E205,\"制造费用\",凭证检查!F6:F205,\"福利费\",凭证检查!M6:M205,\"<>\")";
                                    WorkingPaper.Wb.Worksheets["制造费用、生产成本"].Range["F24"].Formula =
                                        "=-SUMIFS(凭证检查!G6:G205,凭证检查!E6:E205,\"制造费用\",凭证检查!F6:F205,\"职工教育经费\",凭证检查!M6:M205,\"<>\")";
                                    WorkingPaper.Wb.Worksheets["制造费用、生产成本"].Range["F25"].Formula =
                                        "=-SUMIFS(凭证检查!G6:G205,凭证检查!E6:E205,\"制造费用\",凭证检查!F6:F205,\"业务招待费\",凭证检查!M6:M205,\"<>\")";
                                    WorkingPaper.Wb.Worksheets["制造费用、生产成本"].Range["F38"].Formula = "=-F23-F24-F25";

                                    WorkingPaper.Wb.Worksheets["营业费用"].Range["F7"].Formula =
                                        "=-SUMIFS(凭证检查!G6:G205,凭证检查!E6:E205,\"营业费用\",凭证检查!F6:F205,\"福利费\",凭证检查!M6:M205,\"<>\")-SUMIFS(凭证检查!G6:G205,凭证检查!E6:E205,\"销售费用\",凭证检查!F6:F205,\"福利费\",凭证检查!M6:M205,\"<>\")";
                                    WorkingPaper.Wb.Worksheets["营业费用"].Range["F8"].Formula =
                                        "=-SUMIFS(凭证检查!G6:G205,凭证检查!E6:E205,\"营业费用\",凭证检查!F6:F205,\"职工教育经费\",凭证检查!M6:M205,\"<>\")-SUMIFS(凭证检查!G6:G205,凭证检查!E6:E205,\"销售费用\",凭证检查!F6:F205,\"职工教育经费\",凭证检查!M6:M205,\"<>\")";
                                    WorkingPaper.Wb.Worksheets["营业费用"].Range["F10"].Formula =
                                        "=-SUMIFS(凭证检查!G6:G205,凭证检查!E6:E205,\"营业费用\",凭证检查!F6:F205,\"业务招待费\",凭证检查!M6:M205,\"<>\")-SUMIFS(凭证检查!G6:G205,凭证检查!E6:E205,\"销售费用\",凭证检查!F6:F205,\"业务招待费\",凭证检查!M6:M205,\"<>\")";
                                    WorkingPaper.Wb.Worksheets["营业费用"].Range["F42"].Formula = "=-F7-F8-F10";

                                    WorkingPaper.Wb.Worksheets["管理费用"].Range["F7"].Formula =
                                        "=-SUMIFS(凭证检查!G6:G205,凭证检查!E6:E205,\"管理费用\",凭证检查!F6:F205,\"福利费\",凭证检查!M6:M205,\"<>\")";
                                    WorkingPaper.Wb.Worksheets["管理费用"].Range["F8"].Formula =
                                        "=-SUMIFS(凭证检查!G6:G205,凭证检查!E6:E205,\"管理费用\",凭证检查!F6:F205,\"职工教育经费\",凭证检查!M6:M205,\"<>\")";
                                    WorkingPaper.Wb.Worksheets["管理费用"].Range["F10"].Formula =
                                        "=-SUMIFS(凭证检查!G6:G205,凭证检查!E6:E205,\"管理费用\",凭证检查!F6:F205,\"业务招待费\",凭证检查!M6:M205,\"<>\")";
                                    WorkingPaper.Wb.Worksheets["管理费用"].Range["F42"].Formula = "=-F7-F8-F10";

                                    //期间费用
                                    WorkingPaper.Wb.Worksheets["A104000期间费用明细表"].Range["C6:C29"].Replace("营业费用!D", "营业费用!H");
                                    WorkingPaper.Wb.Worksheets["A104000期间费用明细表"].Range["E6:E29"].Replace("管理费用!D", "管理费用!H");
                                    WorkingPaper.Wb.Worksheets["A104000期间费用明细表"].Range["G6:G29"].Replace("财务费用!D", "财务费用!H");

                                    WorkingPaper.Wb.Sheets.Add(After: WorkingPaper.Wb.Worksheets["在建工程审核表"],
                                        Type: AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "\\对外投资.xlsx");

                                    WorkingPaper.Wb.Worksheets["主页"].Hyperlinks.Add(
                                        WorkingPaper.Wb.Worksheets["主页"].Range["H15"],
                                        "#对外投资!A1", Type.Missing, "#对外投资!A1", "对外投资");
                                    Worksheet SH = WorkingPaper.Wb.Worksheets["对外投资"];
                                    SH.Range["C2"].Formula = "=基本情况!B2";
                                    SH.Range["C3"].Formula = "=基本情况!B7";
                                    SH.Range["F2"].Formula = "=基本情况!B12";
                                    SH.Range["F3"].Formula = "=基本情况!B11";
                                    SH.Range["H2"].Formula = "=TEXT(基本情况!B21,\"yyyy-mm-dd\")";
                                    SH.Range["H3"].Formula = "=TEXT(基本情况!B22,\"yyyy-mm-dd\")";
                                    SH.Range["C26"].Formula = "=IF($H$15<>资产负债!$D$6,\"短期投资账载数与报表数相差\"&RMB($H$15-资产负债!$D$6,2)&\"元！\",\"短期投资账载数与报表数相符！\")";
                                    SH.Range["G26"].Formula = "=IF($H$25<>资产负债!$D$21+资产负债!$D$22,\"长期投资账载数与报表数相差\"&RMB($H$25-资产负债!$D$21-资产负债!$D$22,2)&\"元！\",\"长期投资账载数与报表数相符！\")";
                                    SH.Range["D27"].Formula = "=IF(OR(H15<>资产负债!D6,H25<>资产负债!D21+资产负债!D22),\"、E\",\"\")";


                                    WorkingPaper.Wb.Sheets.Add(After: WorkingPaper.Wb.Worksheets["其他应付"],
                                        Type: AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "\\借款.xlsx");

                                    WorkingPaper.Wb.Worksheets["主页"].Hyperlinks.Add(
                                        WorkingPaper.Wb.Worksheets["主页"].Range["I12"],
                                        "#借款!A1", Type.Missing, "#借款!A1", "借款");
                                    SH = WorkingPaper.Wb.Worksheets["借款"];
                                    SH.Range["C2"].Formula = "=基本情况!B2";
                                    SH.Range["C3"].Formula = "=基本情况!B7";
                                    SH.Range["G2"].Formula = "=基本情况!B12";
                                    SH.Range["G3"].Formula = "=基本情况!B11";
                                    SH.Range["I2"].Formula = "=TEXT(基本情况!B21,\"yyyy-mm-dd\")";
                                    SH.Range["I3"].Formula = "=TEXT(基本情况!B22,\"yyyy-mm-dd\")";
                                    SH.Range["C26"].Formula = "=IF($D$15<>资产负债!$H$5,\"短期借款账载数与报表数相差\"&RMB($D$15-资产负债!$H$5,2)&\"元！\",\"短期借款账载数与报表数相符！\")";
                                    SH.Range["H26"].Formula = "=IF($D$25<>资产负债!$H$21,\"长期借款账载数与报表数相差\"&RMB($D$25-资产负债!$H$21,2)&\"元！\",\"长期借款账载数与报表数相符！\")";
                                    SH.Range["D27"].Formula = "=IF(OR(I15<>资产负债!H5,I25<>资产负债!H21),\"、E\",\"\")";

                                    SH = WorkingPaper.Wb.Worksheets["检查表"];
                                    SH.Range["A69:D69"].AutoFill(Destination: SH.Range["A69:D73"]);
                                    SH.Hyperlinks.Add(SH.Range["A70"], "#对外投资!C26", Type.Missing, "#对外投资!C26", "短期投资");
                                    SH.Hyperlinks.Add(SH.Range["A71"], "#对外投资!G26", Type.Missing, "#对外投资!G26", "长期投资");
                                    SH.Hyperlinks.Add(SH.Range["A72"], "#借款!C26", Type.Missing, "#借款!C26", "短期借款");
                                    SH.Hyperlinks.Add(SH.Range["A73"], "#借款!H26", Type.Missing, "#借款!H26", "长期借款");
                                    SH.Range["C70"].Formula = "=对外投资!H15-资产负债!$D$62";
                                    SH.Range["C71"].Formula = "=对外投资!$H$25-资产负债!$D$21-资产负债!$D$22";
                                    SH.Range["C72"].Formula = "=借款!$D$15-资产负债!$H$5";
                                    SH.Range["C73"].Formula = "=借款!$D$25-资产负债!$H$21";

                                    #endregion
                                    Banben = "V20170312-" + Banben.Substring(5);
                                }
                                if (Banben.Substring(0, 9) == "V20170312")
                                {
                                    #region 20170312升级为20170422
                                    //插入研发加计汇总表
                                    Wb.Sheets.Add(After: Wb.Worksheets["A110017居民企业资产（股权）划转特殊性税务处理申报表"],
                                        Type: AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "\\研发加计汇总表.xlsx");
                                    Wb.Worksheets["研发加计扣除归集审核表"].Move(
                                        After: Wb.Worksheets["A110017居民企业资产（股权）划转特殊性税务处理申报表"]);
                                    Wb.Worksheets["研发费用加计扣除优惠审核表"].Range["O15"].Formula = "=研发加计扣除归集审核表!D76";
                                    Wb.Worksheets["研发费用加计扣除优惠审核表"].Range["S15"].Formula = "=研发加计扣除归集审核表!D78";
                                    Wb.Worksheets["研发费用加计扣除优惠审核表"].Range["T15"].Formula = "=研发加计扣除归集审核表!D79";
                                    Wb.Worksheets["主页"].Hyperlinks.Add(
                                        Wb.Worksheets["主页"].Range["D17"],
                                        "#研发加计扣除归集审核表!A1", Type.Missing, "#研发加计扣除归集审核表!A1", "研发加计扣除归集审核表");
                                    Wb.Worksheets["主页"].Range["G21:G28"].ClearContents();
                                    Wb.Worksheets["主页"].Range["G21"].Value2 = "受控外国企业信息报告表";
                                    Wb.Worksheets["主页"].Range["G22"].Value2 = "居民企业资产（股权）划转特殊性税务处理申报表";
                                    Wb.Worksheets["主页"].Range["G23"].Value2 = "非货币性资产投资递延纳税调整明细表";
                                    Wb.Worksheets["主页"].Range["G24"].Value2 = "企业重组所得税特殊性税务处理报告表 ";
                                    Wb.Worksheets["主页"].Hyperlinks.Add(
                                        Wb.Worksheets["主页"].Range["G25"],
                                        "#研发项目可加计扣除研究开发费用情况归集表!A1", Type.Missing, "#研发项目可加计扣除研究开发费用情况归集表!A1", "研发项目可加计扣除研究开发费用情况归集表");
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Rows["49:51"].Delete();
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["A44:B48"].ClearContents();
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B44"].Value2 = "受控外国企业信息报告表";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B45"].Value2 = "居民企业资产（股权）划转特殊性税务处理申报表";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B46"].Value2 = "非货币性资产投资递延纳税调整明细表";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B47"].Value2 = "企业重组所得税特殊性税务处理报告表";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Hyperlinks.Add(
                                        Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B48"],
                                        "#研发项目可加计扣除研究开发费用情况归集表!A1", Type.Missing, "#研发项目可加计扣除研究开发费用情况归集表!A1", "研发项目可加计扣除研究开发费用情况归集表");
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["F48"].Formula = "=IF(研发项目可加计扣除研究开发费用情况归集表!D79<>0,\"是\",\"否\")";
                                    Wb.Worksheets["检查表"].Hyperlinks.Add(
                                        Wb.Worksheets["检查表"].Range["A8"],
                                        "#研发加计扣除归集审核表!A1", Type.Missing, "#研发加计扣除归集审核表!A1", "研发加计扣除归集审核表");

                                    //小微公式修改
                                    Wb.Worksheets["减免所得税优惠审核表"].Range["H4"].Formula = @"=IF(H3<>"""",IF('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!D26<=300000,ROUND('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!D26*0.15,2),""""))";
                                    Wb.Worksheets["减免所得税优惠审核表"].Range["H5"].Formula = @"=IF(H3<>"""",IF('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!D26<=300000,ROUND('A100000中华人民共和国企业所得税年度纳税申报表（A类）'!D26*0.15,2),""""))";


                                    //修改事项说明
                                    Wb.Worksheets["事项说明"].Range["A15"].Formula = "=\"    贵单位\" & IF(OR(基本情况!B6<>\"12月31日\",基本情况!F5<>\"01\",基本情况!G5<>\"01\"),基本情况!B7,基本情况!B4&\"年度\") & \"账面销售（营业）收入\" & RMB(主营收支!H20+其他业务!C18+其他事项!E10,2) & \"元，利润总额\" & RMB(利润!C37,2) & \"元，经审核调整如下：\"";
                                    Wb.Worksheets["事项说明"].Range["D27"].Formula = "=利润!C37+C16-C22";

                                    //A105000纳税调整项目明细表 不征税收入公式修改
                                    Wb.Worksheets["A105000纳税调整项目明细表"].Range["E12"].Formula = "=MAX(其他事项!F8,0)+MAX(A105040专项用途财政性资金纳税调整表!P13,0)";
                                    Wb.Worksheets["A105000纳税调整项目明细表"].Range["F12"].Formula = "=MAX(-其他事项!F8,0)+MAX(A105040专项用途财政性资金纳税调整表!F13,0)";

                                    #endregion
                                    Banben = "V20170422-" + Banben.Substring(5);
                                }

                                if (Banben.Substring(0, 9) == "V20170422")
                                {
                                    #region 20170422升级为20170517

                                    //1、期间费用替换
                                    try
                                    { WorkingPaper.Wb.Worksheets["A104000期间费用明细表"].Range["C6:G29"].Replace("D", "H"); }
                                    finally { }
                                    //2、税收累计折旧
                                    WorkingPaper.Wb.Worksheets["固资折旧"].Range["A8:A12"].FormulaR1C1 =
                                        "=R[10]C[2]+R[10]C[6]-R[10]C[4]";
                                    //3、基本情况（封面） B12 二签身份证号
                                    WorkingPaper.Wb.Worksheets["基本情况（封面）"].Range["B12"].Formula =
                                        "=IFERROR(VLOOKUP(\'基本情况（封面）\'!B13,IF(基本情况!B8=\"中汇百邦（厦门）税务师事务所有限公司\",首页!C:D,首页!E:F),2,0),\"\")";
                                    //4、研发费用加计扣除优惠审核表 去掉O15和S15
                                    WorkingPaper.Wb.Worksheets["研发费用加计扣除优惠审核表"].Range["O15"].Value = 0;
                                    WorkingPaper.Wb.Worksheets["研发费用加计扣除优惠审核表"].Range["S15"].Value = 0;
                                    //5、（三）企业所得税年度纳税申报表填报表单  F31 取数公式
                                    WorkingPaper.Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["F31"].Formula =
                                        "=IF(SUM(A107014研发费用加计扣除优惠明细表!T15)<>0,\"是\",\"否\")";
                                    //6、研发加计扣除归集审核表 D22 取数公式
                                    WorkingPaper.Wb.Worksheets["研发加计扣除归集审核表"].Range["D22"].Formula =
                                        "=SUM(D23:D25)";
                                    //7、研发项目可加计扣除研究开发费用情况归集表 D22 取数公式
                                    WorkingPaper.Wb.Worksheets["研发项目可加计扣除研究开发费用情况归集表"].Range["D22"].Formula =
                                        "=SUM(D23:D25)";
                                    //8、A100000中华人民共和国企业所得税年度纳税申报表（A类）   D10 = 利润!C24 D11 = 利润!C25
                                    WorkingPaper.Wb.Worksheets["A100000中华人民共和国企业所得税年度纳税申报表（A类）"].Range["D10"].Formula =
                                        "=利润!C24";
                                    WorkingPaper.Wb.Worksheets["A100000中华人民共和国企业所得税年度纳税申报表（A类）"].Range["D11"].Formula =
                                        "=利润!C25";
                                    //9、A000000企业基础信息表 从业人数 B8 = IFERROR(ROUNDUP(AVERAGE(INDIRECT("社保明细工资人数!J" & 8 + VALUE(基本情况!F5) & ":J" & 8 + VALUE(基本情况!F6))), 0), 0)                        
                                    WorkingPaper.Wb.Worksheets["A000000企业基础信息表"].Range["B8"].Formula =
                                        "=IFERROR(ROUNDUP(AVERAGE(INDIRECT(\"社保明细工资人数!J\"& 8+VALUE(基本情况!F5) &\":J\" & 8+VALUE(基本情况!F6))),0),0)";
                                    //10、A000000企业基础信息表 资产总额  B9 = ROUND((资产负债!C35 + 资产负债!D35) / 2 / 10000,2)
                                    WorkingPaper.Wb.Worksheets["A000000企业基础信息表"].Range["B9"].Formula =
                                    "=ROUND((资产负债!C35+资产负债!D35)/2/10000,2)";
                                    //11、基本情况 B38 = IF(地税、基本情况!X31 = "", "小企业会计准则", 地税、基本情况!X31)
                                    WorkingPaper.Wb.Worksheets["基本情况"].Range["B38"].Formula =
                                    "=IF(地税、基本情况!X31=\"\",\"小企业会计准则\",地税、基本情况!X31)";

                                    #endregion
                                    Banben = "V20170517-" + Banben.Substring(5);
                                }

                                if (Banben.Substring(0, 9) == "V20170517")
                                {
                                    #region 20170517升级为20171222
                                    //添加合并表格
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Rows["44:51"].Insert(Shift: XlInsertShiftDirection.xlShiftDown);
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["C44:C51"].FormulaR1C1 = "=RC[3]";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["A44"].Value2 = "A110010";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B44"].Value2 = "    特殊性处理报告表";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["A45"].Value2 = "A110011";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B45"].Value2 = "    债务重组报告表";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["A46"].Value2 = "A110012";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B46"].Value2 = "    股权收购报告表";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["A47"].Value2 = "A110013";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B47"].Value2 = "    资产收购报告表";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["A48"].Value2 = "A110014";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B48"].Value2 = "    企业合并报告表";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["A49"].Value2 = "A110015";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B49"].Value2 = "    企业分立报告表";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["A50"].Value2 = "A110016";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B50"].Value2 = "    非货币资产投资递延纳税调整表";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["A51"].Value2 = "A110017";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B51"].Value2 = "    居民企业资产（股权）划转特殊性税务处理申报表";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Hyperlinks.Add(
                                        Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B44"], "#A110010特殊性处理报告表!A1",
                                        Type.Missing, "#A110010特殊性处理报告表!A1", "    特殊性处理报告表");
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Hyperlinks.Add(
                                        Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B45"], "#A110011债务重组报告表!A1",
                                        Type.Missing, "#A110011债务重组报告表!A1", "    债务重组报告表");
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Hyperlinks.Add(
                                        Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B46"], "#'A110012股权收购报告表 '!A1",
                                        Type.Missing, "#'A110012股权收购报告表 '!A1", "    股权收购报告表");
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Hyperlinks.Add(
                                        Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B47"], "#A110013资产收购报告表!A1",
                                        Type.Missing, "#A110013资产收购报告表!A1", "    资产收购报告表");
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Hyperlinks.Add(
                                        Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B48"], "#'A110014企业合并报告表 '!A1",
                                        Type.Missing, "#'A110014企业合并报告表 '!A1", "    企业合并报告表");
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Hyperlinks.Add(
                                        Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B49"], "#A110015企业分立申报表!A1",
                                        Type.Missing, "#A110015企业分立申报表!A1", "    企业分立报告表");
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Hyperlinks.Add(
                                        Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B50"], "#A110016非货币资产投资递延纳税调整表!A1",
                                        Type.Missing, "#A110016非货币资产投资递延纳税调整表!A1", "    非货币资产投资递延纳税调整表 ");
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Hyperlinks.Add(
                                        Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["B51"],
                                        "#'A110017居民企业资产（股权）划转特殊性税务处理申报表'!A1", Type.Missing,
                                        "#'A110017居民企业资产（股权）划转特殊性税务处理申报表'!A1", "    居民企业资产（股权）划转特殊性税务处理申报表");
                                    Wb.Worksheets["主页"].Hyperlinks.Add(Wb.Worksheets["主页"].Range["G26"],
                                        "#A110010特殊性处理报告表!A1", Type.Missing, "#A110010特殊性处理报告表!A1", "特殊性处理报告表");
                                    Wb.Worksheets["主页"].Hyperlinks.Add(Wb.Worksheets["主页"].Range["G27"],
                                        "#A110011债务重组报告表!A1", Type.Missing, "#A110011债务重组报告表!A1", "债务重组报告表");
                                    Wb.Worksheets["主页"].Hyperlinks.Add(Wb.Worksheets["主页"].Range["G28"],
                                        "#'A110012股权收购报告表 '!A1", Type.Missing, "#'A110012股权收购报告表 '!A1", "股权收购报告表");
                                    Wb.Worksheets["主页"].Hyperlinks.Add(Wb.Worksheets["主页"].Range["G29"],
                                        "#A110013资产收购报告表!A1", Type.Missing, "#A110013资产收购报告表!A1", "资产收购报告表");
                                    Wb.Worksheets["主页"].Hyperlinks.Add(Wb.Worksheets["主页"].Range["G30"],
                                        "#'A110014企业合并报告表 '!A1", Type.Missing, "#'A110014企业合并报告表 '!A1", "企业合并报告表");
                                    Wb.Worksheets["主页"].Hyperlinks.Add(Wb.Worksheets["主页"].Range["G31"],
                                        "#A110015企业分立申报表!A1", Type.Missing, "#A110015企业分立申报表!A1", "企业分立报告表");
                                    Wb.Worksheets["主页"].Hyperlinks.Add(Wb.Worksheets["主页"].Range["G32"],
                                        "#A110016非货币资产投资递延纳税调整表!A1", Type.Missing, "#A110016非货币资产投资递延纳税调整表!A1",
                                        "非货币资产投资递延纳税调整表");
                                    Wb.Worksheets["主页"].Hyperlinks.Add(Wb.Worksheets["主页"].Range["G33"],
                                        "#'A110017居民企业资产（股权）划转特殊性税务处理申报表'!A1", Type.Missing,
                                        "#'A110017居民企业资产（股权）划转特殊性税务处理申报表'!A1", "居民企业资产（股权）划转特殊性税务处理申报表");


                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["F44"].Value2 = "否";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["F45"].Value2 = "否";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["F46"].Value2 = "否";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["F47"].Value2 = "否";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["F48"].Value2 = "否";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["F49"].Value2 = "否";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["F50"].Value2 = "否";
                                    Wb.Worksheets["（三）企业所得税年度纳税申报表填报表单"].Range["F51"].Value2 = "否";

                                    #endregion
                                    Banben = "V20171222-" + Banben.Substring(5);
                                }

                                WorkingPaper.Wb.Worksheets["首页"].Range["A1"].Value2 = Banben;
                                WorkingPaper.Wb.Worksheets["首页"].Protect();
                                Globals.WPToolAddln.Application.StatusBar = false;
                                MessageBox.Show("升级完成，请检查！");
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("当前版本为：" + Banben + "，最新版本为：V20171222。不需要升级", "提示！",
                            MessageBoxButtons.OK);
                    }
                }

            }

        }

        private void btn工具设置_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            if(WorkingPaper.OOO)
            {
                WorkingPaper.Wb.Application.ScreenUpdating = true;
                Globals.WPToolAddln.Application.Workbooks.Open(
                    AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "\\打印报告.xlsx",
                    XlUpdateLinks.xlUpdateLinksNever);
                MessageBox.Show("如果打印报告权限出错，请打开打印报告源文件，修改权限为可编辑！");
            }
        }

        private void btn底稿打印_Click(object sender, RibbonControlEventArgs e)
        {
            if (WorkingPaper.版本号 != 2018)
            {

                底稿打印 dgdy = new 底稿打印();
                if (dgdy.ShowDialog() == DialogResult.Yes)
                {
                    WorkingPaper.Wb.PrintPreview();
                }
            }
        }

        private void btn打印报告_Click(object sender, RibbonControlEventArgs e)
        {
            if (WorkingPaper.OOO)
            {
                if (WorkingPaper.版本号 != 2018)
                {

                    if (Math.Round(CU.Shuzi(WorkingPaper.Wb.Worksheets["A107040减免所得税优惠明细表"].Range["D7"].Value2) +
                                   CU.Shuzi(WorkingPaper.Wb.Worksheets["A107040减免所得税优惠明细表"].Range["D8"].Value2), 2) !=
                        Math.Round(CU.Shuzi(WorkingPaper.Wb.Worksheets["A107040减免所得税优惠明细表"].Range["D6"].Value2), 2))
                    {
                        MessageBox.Show("A107040减免所得税优惠明细表，D6不等于D7+D8，请检查后重试。");
                        return;
                    }

                    if (MessageBox.Show("现在要切换到打印状态。是否继续？", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        string 打印文件路径 = WorkingPaper.Wb.Path + "\\" + Wb.Name.Substring(0, Wb.Name.LastIndexOf(".")) +
                                        "打印报告.xlsx";
                        try
                        {
                            Globals.WPToolAddln.Application.StatusBar = "正在导出报告...";
                            Globals.WPToolAddln.Application.DisplayAlerts = false;
                            File.Copy(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "\\打印报告.xlsx", 打印文件路径,
                                true);
                            CU.事项说明();
                            WorkingPaper.wb打印 =
                                Globals.WPToolAddln.Application.Workbooks.Open(打印文件路径,
                                    XlUpdateLinks.xlUpdateLinksNever);
                            Globals.WPToolAddln.Application.ScreenUpdating = false;
                            WorkingPaper.wb打印.ChangeLink(Name: @"E:\税审底稿 模板.xlsx", NewName: Wb.FullName,
                                Type: XlLinkType.xlLinkTypeExcelLinks);
                            //Newbook.UpdateLink(WorkingPaper.Wb.FullName, XlLinkType.xlLinkTypeExcelLinks);
                            //WorkingPaper.wb打印.BreakLink(WorkingPaper.Wb.FullName, XlLinkType.xlLinkTypeExcelLinks);
                            CU.自动调整行高("企业基本情况", "C10:F10", 46.78);
                            CU.自动调整行高("企业基本情况", "A128:F128", 85.22);
                            CU.自动调整行高("A000000企业基础信息表", "B7", 15.67);
                            CU.自动调整行高("A000000企业基础信息表", "A21", 18.44);
                            CU.自动调整行高("A000000企业基础信息表", "A22", 18.44);
                            CU.自动调整行高("A000000企业基础信息表", "A23", 18.44);
                            CU.自动调整行高("A000000企业基础信息表", "A24", 18.44);
                            CU.自动调整行高("A000000企业基础信息表", "A25", 18.44);
                            CU.自动调整行高("A000000企业基础信息表", "A28", 18.44);
                            CU.自动调整行高("A000000企业基础信息表", "A29", 18.44);
                            CU.自动调整行高("A000000企业基础信息表", "A30", 18.44);
                            CU.自动调整行高("A000000企业基础信息表", "A31", 18.44);
                            CU.自动调整行高("A000000企业基础信息表", "A32", 18.44);
                            WorkingPaper.wb打印.Sheets["企业基本情况"].Range["$H$21:$H$128"]
                                .AutoFilter(Field: 1, Criteria1: "=1");
                            object[,] 表单 = WorkingPaper.wb打印.Sheets["（三）企业所得税年度纳税申报表填报表单"].Range["$C$3:$D$56"].Value2;
                            for (int i = 1; i <= 54; i++)
                            {
                                if (CU.Zifu(表单[i, 1]) == "否" && CU.Zifu(表单[i, 2]) != "")
                                {
                                    WorkingPaper.wb打印.Sheets[CU.Zifu(表单[i, 2])].Visible = false;
                                }

                            }

                            if (CU.Zifu(表单[54, 1]) == "是")
                            {
                                object[,] 其他相关费用 = WorkingPaper.wb打印.Sheets["研发项目可加计扣除研究开发费用情况归集表"].Range["$B$35:$B$71"]
                                    .Value2;
                                Boolean konghang = false;
                                int i;
                                for (i = 1; i <= 37; i++)
                                {
                                    if (CU.Zifu(其他相关费用[i, 1]) == "")
                                    {
                                        konghang = true;
                                        break;
                                    }
                                }

                                if (konghang)
                                    WorkingPaper.wb打印.Sheets["研发项目可加计扣除研究开发费用情况归集表"].Rows[(i + 34).ToString() + ":71"]
                                        .Hidden = true;
                            }

                            if (WorkingPaper.Wb.Worksheets["基本情况"].range("B8").value == "厦门明正税务师事务所有限公司")
                            {
                                WorkingPaper.wb打印.Sheets["中汇封面"].Visible = false;
                            }
                            else
                            {
                                WorkingPaper.wb打印.Sheets["明正封面"].Visible = false;
                            }

                            if (CU.Zifu(WorkingPaper.wb打印.Sheets["A109010企业所得税汇总纳税分支机构所得税分配表"].Range["C3"].Value2) ==
                                "分支机构")
                            {
                                WorkingPaper.wb打印.Sheets["分支机构企业所得税申报表（A类）"].Visible = true;
                            }

                            List<string> lists = new List<string>();

                            int C = WorkingPaper.wb打印.Worksheets.Count;
                            for (int i = 1; i <= C; i++)
                            {
                                //MessageBox.Show(WorkingPaper.wb打印.Worksheets[i].Visible.ToString()); 
                                if (WorkingPaper.wb打印.Sheets[i].Visible == -1)
                                {
                                    lists.Add(WorkingPaper.wb打印.Worksheets[i].Name);
                                }
                            }

                            string[] s = lists.ToArray();

                            WorkingPaper.wb打印.Worksheets[s].Select();
                            Globals.WPToolAddln.Application.DisplayAlerts = true;
                            Globals.WPToolAddln.Application.ScreenUpdating = true;
                            Globals.WPToolAddln.Application.StatusBar = false;
                            WorkingPaper.wb打印.Activate();
                            WorkingPaper.wb打印.PrintPreview();
                            //Newbook.Save();
                            //Newbook.Close();
                            WorkingPaper.wb打印 = null;
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

        private void Application_SheetActivate(object sh)
        {
            if(OOO)
            {
                string ss = Wb.ActiveSheet.Name;
                Contents con;
                if (Excel版本 == 13)
                {
                    int hwnd = Globals.WPToolAddln.Application.ActiveWindow.Hwnd;
                    Cons.TryGetValue(hwnd, out con);
                }
                else
                    con = Excel10Con;
                if (con != null)
                {
                    switch (ss)
                    {
                        case "余额表":
                        case "税金申报明细":
                        case "基本情况":
                            con.显示选项卡(ss);
                            break;
                        case "检查表":
                            Globals.WPToolAddln.Application.SheetFollowHyperlink += Application_SheetFollowHyperlink;
                            break;
                        case "A000000企业基础信息表":
                            if (WorkingPaper.版本号!=2018)
                            {
                                Globals.WPToolAddln.Application.SheetSelectionChange += Application_SheetSelectionChange;
                            }
                            
                            break;
                        case "固资折旧":
                            object[,] 期末原值 = Wb.ActiveSheet.Range["F8:F12"].Value2;
                            object[,] 期末折旧 = Wb.ActiveSheet.Range["F18:F22"].Value2;
                            object[,] 期末税收折旧 = Wb.ActiveSheet.Range["A8:A12"].Value2;
                            if(CU.Shuzi(期末原值[1,1])<CU.Shuzi(期末折旧[1,1]) || CU.Shuzi(期末原值[1, 1]) < CU.Shuzi(期末税收折旧[1, 1]))
                            {
                                MessageBox.Show("房屋建筑累计折旧大于原值！");
                            }
                            if (CU.Shuzi(期末原值[2, 1]) < CU.Shuzi(期末折旧[2, 1]) || CU.Shuzi(期末原值[2, 1]) < CU.Shuzi(期末税收折旧[2, 1]))
                            {
                                MessageBox.Show("机械设备累计折旧大于原值！");
                            }

                            if (CU.Shuzi(期末原值[3, 1]) < CU.Shuzi(期末折旧[3, 1]) || CU.Shuzi(期末原值[3, 1]) < CU.Shuzi(期末税收折旧[3, 1]))
                            {
                                MessageBox.Show("工器家具累计折旧大于原值！");
                            }
                            if (CU.Shuzi(期末原值[4, 1]) < CU.Shuzi(期末折旧[4, 1]) || CU.Shuzi(期末原值[4, 1]) < CU.Shuzi(期末税收折旧[4, 1]))
                            {
                                MessageBox.Show("运输工具累计折旧大于原值！");
                            }
                            if (CU.Shuzi(期末原值[5, 1]) < CU.Shuzi(期末折旧[5, 1]) || CU.Shuzi(期末原值[5, 1]) < CU.Shuzi(期末税收折旧[5, 1]))
                            {
                                MessageBox.Show("电子设备累计折旧大于原值！");
                            }
                            break;
                        default:
                            Globals.WPToolAddln.Application.SheetFollowHyperlink -= Application_SheetFollowHyperlink;
                            Globals.WPToolAddln.Application.SheetSelectionChange -= Application_SheetSelectionChange;
                            con.显示选项卡("");
                            break;
                    }
                }
            }
        }

        private void Application_SheetSelectionChange(object sh, Range target)
        {
            if (target.Address == "$B$15:$F$15")
            {
                存货计价 ch= new 存货计价();
                
                ch.ShowDialog();
            }
        }
        

        private void Application_SheetFollowHyperlink(object sh, Hyperlink target)
        {
            if (WorkingPaper.OOO)
            {
                try
                {
                    Globals.WPToolAddln.Application.ScreenUpdating = false;
                    string add = target.Range.Hyperlinks[1].SubAddress;
                    add = add.Substring(0, add.IndexOf("!")).Replace("'", "");
                    
                    Wb.Worksheets[add].Visible = true;
                    if (Wb.ActiveSheet.Cells[1, 7].Value.ToString() == "跳转超链接所选页面")
                        Wb.Worksheets[add].Select();

                    Globals.WPToolAddln.Application.ScreenUpdating = true;
                }
                catch (Exception)
                {
                    Globals.WPToolAddln.Application.ScreenUpdating = true;
                }
            }
        }

        private void btnUpdata_Click(object sender, RibbonControlEventArgs e)
        {
            更新(true);
        }

        private void btnGetURL_Click(object sender, RibbonControlEventArgs e)
        {
            string 下载地址 = Contents.获取版本号("https://zhuanlan.zhihu.com/p/26527380");
            if (下载地址 == "获取失败")
                MessageBox.Show("版本获取失败，请检查网络后重试！");
            else if (MessageBox.Show("最新版本下载地址为：" + 下载地址 + "，是否用默认浏览器打开？", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(下载地址);
            }
            
        }

        private void 更新(Boolean O)
        {
            string 最新版本 = Contents.获取版本号("https://zhuanlan.zhihu.com/p/26474507");
            if (最新版本 == "获取失败")
                MessageBox.Show("版本获取失败，请检查网络后重试！");
            else
                if (最新版本 != 当前版本)
            {
                MessageBox.Show("当前版本为：" + 当前版本 + "，发现新版本：" + 最新版本 + "。请通过微信公众号下载新版本！");
            }
            else
                if(O)
                MessageBox.Show("当前版本为：" + 当前版本 + "，最新版本为：" + 最新版本 + "，不需要更新。请关注微信公众号以获取最新版本动态！");
        }

        //工作簿激活事件
        private void btnGongzhonghao_Click(object sender, RibbonControlEventArgs e)
        {
            Contact tac = new Contact();
            tac.ShowDialog();
        }

        private void Application_WorkbookActivate(Workbook wb)
        {

            if (CU.文件判断())
            {
                
                Contents con = new Contents();
                if (Excel版本 == 13)
                {
                    int hwnd = Globals.WPToolAddln.Application.ActiveWindow.Hwnd;
                    Microsoft.Office.Tools.CustomTaskPane mypane;
                    TaskPanels.TryGetValue(hwnd, out mypane);
                    if (mypane != null)
                    {
                        tb显示目录.Checked = mypane.Visible;
                    }
                    else
                    {
                        Microsoft.Office.Tools.CustomTaskPane pane = Globals.WPToolAddln.CustomTaskPanes.Add(con,
                            "税审底稿工具",
                            Globals.WPToolAddln.Application.ActiveWindow);
                        //这一步很重要将决定是否显示到当前窗口，第三个参数的意思就是依附到那个窗口
                        //pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                        pane.Width = 300;
                        TaskPanels.Add(hwnd, pane);
                        Cons.Add(hwnd, con);
                        pane.VisibleChanged += new EventHandler(MyTaskpane_VisibleChanged);
                        pane.Visible = tb显示目录.Checked;
                    }
                    Cons.TryGetValue(hwnd, out con);
                }
                else
                {
                    con = Excel10Con;
                    Excel10Taskpane.Visible = true;
                    tb显示目录.Checked = true;
                }
                Globals.WPToolAddln.Application.SheetActivate += Application_SheetActivate;
                string ss = wb.ActiveSheet.Name;
                if (con != null)
                {
                    switch (ss)
                    {
                        case "余额表":
                        case "税金申报明细":
                        case "基本情况":
                            con.显示选项卡(ss);
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
                            con.显示选项卡("");
                            break;
                    }
                }
            }
            else
            {
                if(Excel版本==10|| Excel版本 == 07)
                    if (Excel10Taskpane != null) Excel10Taskpane.Visible = false;
                tb显示目录.Checked = false;
                Globals.WPToolAddln.Application.SheetActivate -= Application_SheetActivate;
                Globals.WPToolAddln.Application.SheetFollowHyperlink -= Application_SheetFollowHyperlink;
            }
            添加右键();
        }

        #region 导出功能

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

        private void 导出PDF(object sender, RibbonControlEventArgs e)
        {
            if (MessageBox.Show("现在将当前可见工作表导出为PDF。是否继续？", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                SaveFileDialog Sv = new SaveFileDialog();
                Sv.Filter = "PDF文件(*.pdf)|*.pdf";
                Sv.FileName = "税审工作表导出";
                Sv.Title = "导出当前可见工作表";
                Sv.OverwritePrompt = true;
                Sv.InitialDirectory = Globals.WPToolAddln.Application.ActiveWorkbook.Path;
                if (Sv.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        Globals.WPToolAddln.Application.StatusBar = "正在导出可见工作表...";
                        Globals.WPToolAddln.Application.DisplayAlerts = false;
                        Globals.WPToolAddln.Application.ScreenUpdating = false;
                        Globals.WPToolAddln.Application.ActiveWorkbook.ExportAsFixedFormat(
                            Type: XlFixedFormatType.xlTypePDF,
                            Filename: Sv.FileName.ToString(), IgnorePrintAreas: false, OpenAfterPublish: true);
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
                        if (Excel版本 == 07)
                        {
                            MessageBox.Show("当前Excel为2007版本，建议安装 SaveAsPDFandXPS 后重试一下。");
                        }
                    }
                }
            }
        }

        private void 导出成03(object sender, RibbonControlEventArgs e)
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

        #endregion
    }
}
