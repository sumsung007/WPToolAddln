using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace 百邦所得税汇算底稿工具
{
    public partial class 底稿打印 : Form
    {

        string[] MustName, MustArea, MustDirection, MustZoom, ChooseName, ChooseArea, ChooseCondition, 
            ChooseDirection, ChooseZoom, NonName;
        string group;
        private int DYYear;

        private void Lv待选_MouseDoubleClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (lv待选.SelectedItems.Count > 0 && e.Button==MouseButtons.Left)
            {
                this.lv选中.BeginUpdate();
                this.lv待选.BeginUpdate();
                foreach (ListViewItem lvi in lv待选.SelectedItems)
                {
                    group = lvi.Group.Name;
                    lv待选.Items.Remove(lvi);
                    switch (group)
                    {
                        case "MustGroup":
                            lvi.Group = lv选中.Groups["MustGroup"];
                            break;
                        case "ChooseGroup":
                            lvi.Group = lv选中.Groups["ChooseGroup"];
                            break;
                        case "NonGroup":
                            lvi.Group = lv选中.Groups["NonGroup"];
                            break;
                    }
                    lv选中.Items.Add(lvi);
                }
                this.lv选中.EndUpdate();
                this.lv待选.EndUpdate();
            }
        }
        private void Lv选中_MouseDoubleClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (lv选中.SelectedItems.Count > 0 && e.Button == MouseButtons.Left)
            {
                this.lv选中.BeginUpdate();
                this.lv待选.BeginUpdate();
                foreach (ListViewItem lvi in lv选中.SelectedItems)
                {
                    group = lvi.Group.Name;
                    lv选中.Items.Remove(lvi);
                    switch (group)
                    {
                        case "MustGroup":
                            lvi.Group = lv待选.Groups["MustGroup"];
                            break;
                        case "ChooseGroup":
                            lvi.Group = lv待选.Groups["ChooseGroup"];
                            break;
                        case "NonGroup":
                            lvi.Group = lv待选.Groups["NonGroup"];
                            break;
                    }
                    lv待选.Items.Add(lvi);
                }
                this.lv选中.EndUpdate();
                this.lv待选.EndUpdate();
            }

        }

        private void btn取消_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn全移_Click(object sender, EventArgs e)
        {
            if (lv选中.Items.Count > 0)
            {
                this.lv选中.BeginUpdate();
                this.lv待选.BeginUpdate();
                foreach (ListViewItem lvi in lv选中.Items)
                {
                    group = lvi.Group.Name;
                    lv选中.Items.Remove(lvi);
                    switch (group)
                    {
                        case "MustGroup":
                            lvi.Group = lv待选.Groups["MustGroup"];
                            break;
                        case "ChooseGroup":
                            lvi.Group = lv待选.Groups["ChooseGroup"];
                            break;
                        case "NonGroup":
                            lvi.Group = lv待选.Groups["NonGroup"];
                            break;
                    }
                    lv待选.Items.Add(lvi);
                }
                this.lv选中.EndUpdate();
                this.lv待选.EndUpdate();
            }

        }

        private void btn全选_Click(object sender, EventArgs e)
        {
            if (lv待选.Items.Count > 0)
            {
                this.lv选中.BeginUpdate();
                this.lv待选.BeginUpdate();
                foreach (ListViewItem lvi in lv待选.Items)
                {
                    group = lvi.Group.Name;
                    lv待选.Items.Remove(lvi);
                    switch (group)
                    {
                        case "MustGroup":
                            lvi.Group = lv选中.Groups["MustGroup"];
                            break;
                        case "ChooseGroup":
                            lvi.Group = lv选中.Groups["ChooseGroup"];
                            break;
                        case "NonGroup":
                            lvi.Group = lv选中.Groups["NonGroup"];
                            break;
                    }
                    lv选中.Items.Add(lvi);
                }
                this.lv选中.EndUpdate();
                this.lv待选.EndUpdate();
            }

        }

        private void btn移出_Click(object sender, EventArgs e)
        {
            if (lv选中.SelectedItems.Count > 0)
            {
                this.lv选中.BeginUpdate();
                this.lv待选.BeginUpdate();
                foreach (ListViewItem lvi in lv选中.SelectedItems)
                {
                    group = lvi.Group.Name;
                    lv选中.Items.Remove(lvi);
                    switch (group)
                    {
                        case "MustGroup":
                            lvi.Group = lv待选.Groups["MustGroup"];
                            break;
                        case "ChooseGroup":
                            lvi.Group = lv待选.Groups["ChooseGroup"];
                            break;
                        case "NonGroup":
                            lvi.Group = lv待选.Groups["NonGroup"];
                            break;
                    }
                    lv待选.Items.Add(lvi);
                }
                this.lv选中.EndUpdate();
                this.lv待选.EndUpdate();
            }

        }

        private void btn选中_Click(object sender, EventArgs e)
        {
            if (lv待选.SelectedItems.Count > 0)
            {
                this.lv选中.BeginUpdate();
                this.lv待选.BeginUpdate();
                foreach (ListViewItem lvi in lv待选.SelectedItems)
                {
                    group = lvi.Group.Name;
                    lv待选.Items.Remove(lvi);
                    switch(group)
                    {
                        case "MustGroup":
                            lvi.Group = lv选中.Groups["MustGroup"];
                            break;
                        case "ChooseGroup":
                            lvi.Group = lv选中.Groups["ChooseGroup"];
                            break;
                        case "NonGroup":
                            lvi.Group = lv选中.Groups["NonGroup"];
                            break;
                    }
                    lv选中.Items.Add(lvi);
                }
                this.lv选中.EndUpdate();
                this.lv待选.EndUpdate();
            }
        }

        public 底稿打印(int Year)
        {
            InitializeComponent();
            DYYear = Year;
            刷新();
            lv待选.MouseDoubleClick += Lv待选_MouseDoubleClick;
            lv选中.MouseDoubleClick += Lv选中_MouseDoubleClick;
        }
        
        public void 刷新()
        {
            this.lv选中.BeginUpdate();
            this.lv待选.BeginUpdate();
            this.lv选中.Items.Clear();
            this.lv待选.Items.Clear();
            if (DYYear == 2017)
            {
                //必选打印
                if (Properties.Settings.Default.MustN != "" ||
                    Properties.Settings.Default.MustA != "" ||
                    Properties.Settings.Default.MustD != "" ||
                    Properties.Settings.Default.MustZ != "")
                {
                    MustName = Properties.Settings.Default.MustN.Split(new char[] {'/'});
                    MustArea = Properties.Settings.Default.MustA.Split(new char[] {'/'});
                    MustDirection = Properties.Settings.Default.MustD.Split(new char[] {'/'});
                    MustZoom = Properties.Settings.Default.MustZ.Split(new char[] {'/'});
                    for (int i = 0; i < MustName.Length; i++)
                    {
                        ListViewItem lvi = new ListViewItem();
                        lvi.Group = lv选中.Groups["MustGroup"];
                        lvi.Text = MustName[i];
                        lvi.SubItems.Add("必选");
                        lvi.SubItems.Add(i.ToString());
                        this.lv选中.Items.Add(lvi);
                    }
                }

                //选择打印
                if (Properties.Settings.Default.ChooseN != "" ||
                    Properties.Settings.Default.ChooseA != "" ||
                    Properties.Settings.Default.ChooseC != "" ||
                    Properties.Settings.Default.ChooseD != "" ||
                    Properties.Settings.Default.ChooseZ != "")
                {
                    ChooseName = Properties.Settings.Default.ChooseN.Split(new char[] {'/'});
                    ChooseArea = Properties.Settings.Default.ChooseA.Split(new char[] {'/'});
                    ChooseCondition = Properties.Settings.Default.ChooseC.Split(new char[] {'/'});
                    ChooseDirection = Properties.Settings.Default.ChooseD.Split(new char[] {'/'});
                    ChooseZoom = Properties.Settings.Default.ChooseZ.Split(new char[] {'/'});
                    string[,] Condition = new string[ChooseName.Length, 1];
                    for (int i = 0; i < ChooseName.Length; i++)
                    {
                        Condition[i, 0] = ChooseCondition[i];
                    }

                    WorkingPaper.Wb.Worksheets["首页"].Unprotect();
                    WorkingPaper.Wb.Worksheets["首页"].Range["M1:M" + ChooseName.Length].FormulaArray = Condition;
                    WorkingPaper.Wb.Worksheets["首页"].Protect();
                    object[,] V = WorkingPaper.Wb.Worksheets["首页"].Range["M1:M" + ChooseName.Length].Value2;
                    for (int i = 0; i < ChooseName.Length; i++)
                    {
                        if (CU.Shuzi(V[i + 1, 1]) == 0)
                        {
                            ListViewItem lvi = new ListViewItem();
                            lvi.Group = lv待选.Groups["ChooseGroup"];
                            lvi.Text = ChooseName[i];
                            lvi.SubItems.Add("无数");
                            lvi.SubItems.Add(i.ToString());
                            this.lv待选.Items.Add(lvi);

                        }
                        else
                        {
                            ListViewItem lvi = new ListViewItem();
                            lvi.Group = lv选中.Groups["ChooseGroup"];
                            lvi.Text = ChooseName[i];
                            lvi.SubItems.Add("有数");
                            lvi.SubItems.Add(i.ToString());
                            this.lv选中.Items.Add(lvi);
                        }
                    }
                }

                //不用打印
                if (Properties.Settings.Default.NonN != "")
                {
                    NonName = Properties.Settings.Default.NonN.Split(new char[] {'/'});
                    for (int i = 0; i < NonName.Length; i++)
                    {
                        ListViewItem lvi = new ListViewItem();
                        lvi.Group = lv待选.Groups["NonGroup"];
                        lvi.Text = NonName[i];
                        lvi.SubItems.Add("无需");
                        lvi.SubItems.Add(i.ToString());
                        this.lv待选.Items.Add(lvi);
                    }
                }
            }
            else
            {

                WorkingPaper.Wb.Worksheets["补亏"].PageSetup.PrintComments = XlPrintLocation.xlPrintNoComments;
                string jsonfile = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "\\PrintSetting2018.json";

                using (System.IO.StreamReader file = System.IO.File.OpenText(jsonfile))
                {
                    using (JsonTextReader reader = new JsonTextReader(file))
                    {
                        JObject o = (JObject) JToken.ReadFrom(reader);

                        //必选打印
                        var MustPrintArea = o["MustPrint"];
                        MustName = MustPrintArea.Select(c => c["Sheetname"].ToString()).ToArray();
                        MustArea = MustPrintArea.Select(c => c["Printarea"].ToString()).ToArray();
                        MustDirection = MustPrintArea.Select(c => c["Pagedirection"].ToString()).ToArray();
                        MustZoom = MustPrintArea.Select(c => c["Zoom"].ToString()).ToArray();
                        for (int i = 0; i < MustName.Length; i++)
                        {
                            ListViewItem lvi = new ListViewItem();
                            lvi.Group = lv选中.Groups["MustGroup"];
                            lvi.Text = MustName[i];
                            lvi.SubItems.Add("必选");
                            lvi.SubItems.Add(i.ToString());
                            this.lv选中.Items.Add(lvi);
                        }

                        //选择打印
                        var ChoosePrintArea = o["ChoosePrint"];
                        ChooseName = ChoosePrintArea.Select(c => c["Sheetname"].ToString()).ToArray();
                        ChooseArea = ChoosePrintArea.Select(c => c["Printarea"].ToString()).ToArray();
                        ChooseCondition = ChoosePrintArea.Select(c => c["Formula"].ToString()).ToArray();
                        ChooseDirection = ChoosePrintArea.Select(c => c["Pagedirection"].ToString()).ToArray();
                        ChooseZoom = ChoosePrintArea.Select(c => c["Zoom"].ToString()).ToArray();
                        string[,] Condition = new string[ChooseName.Length, 1];
                        for (int i = 0; i < ChooseName.Length; i++)
                        {
                            Condition[i, 0] = ChooseCondition[i];
                        }

                        WorkingPaper.Wb.Worksheets["辅助表"].Unprotect();
                        WorkingPaper.Wb.Worksheets["辅助表"].Range["X1:X" + ChooseName.Length].FormulaArray = Condition;
                        WorkingPaper.Wb.Worksheets["辅助表"].Protect();
                        object[,] V = WorkingPaper.Wb.Worksheets["辅助表"].Range["X1:X" + ChooseName.Length].Value2;
                        for (int i = 0; i < ChooseName.Length; i++)
                        {
                            if (CU.Shuzi(V[i + 1, 1]) == 0)
                            {
                                ListViewItem lvi = new ListViewItem();
                                lvi.Group = lv待选.Groups["ChooseGroup"];
                                lvi.Text = ChooseName[i];
                                lvi.SubItems.Add("无数");
                                lvi.SubItems.Add(i.ToString());
                                this.lv待选.Items.Add(lvi);

                            }
                            else
                            {
                                ListViewItem lvi = new ListViewItem();
                                lvi.Group = lv选中.Groups["ChooseGroup"];
                                lvi.Text = ChooseName[i];
                                lvi.SubItems.Add("有数");
                                lvi.SubItems.Add(i.ToString());
                                this.lv选中.Items.Add(lvi);
                            }
                        }


                        //不用打印
                        var NoPrintArea = o["NoPrint"];
                        NonName = NoPrintArea.Select(c => c.ToString()).ToArray();
                        for (int i = 0; i < NonName.Length; i++)
                        {
                            ListViewItem lvi = new ListViewItem();
                            lvi.Group = lv待选.Groups["NonGroup"];
                            lvi.Text = NonName[i];
                            lvi.SubItems.Add("无需");
                            lvi.SubItems.Add(i.ToString());
                            this.lv待选.Items.Add(lvi);
                        }
                    }
                }

            }

            this.lv选中.EndUpdate();
            this.lv待选.EndUpdate();
            this.Refresh();
        }

        private void btn打印_Click(object sender, EventArgs e)
        {
            /*Globals.WPToolAddln.Application.ActiveWorkbook.ActiveSheet.Range["A1:D10"].Copy();
            Image img;
            if (System.Windows.Forms.Clipboard.ContainsImage())
            {
                img = System.Windows.Forms.Clipboard.GetImage();
                pictureBox1.Image = img;

            }*/
            Boolean o = Globals.WPToolAddln.Application.Version.ToString() == "12.0" ? false : true;
            try
            {
            string[] HW;
            int j;
            Globals.WPToolAddln.Application.ScreenUpdating = false;

            List<string> lists = new List<string>();
            int n = lv选中.Items.Count;
            for (int i = 0; i < n; i++)
            {
                    //处理Item 
                    if (o) Globals.WPToolAddln.Application.PrintCommunication = false;
                label1.Text = "STEP.2  正在设置打印区域..." + (i + 1).ToString() + "/" + n.ToString();
                this.Refresh();
                string iName = lv选中.Items[i].SubItems[0].Text;
                lists.Add(iName);
                int iNo = Convert.ToInt16(lv选中.Items[i].SubItems[2].Text);
                WorkingPaper.Wb.Worksheets[iName].PageSetup.BlackAndWhite = true;
                WorkingPaper.Wb.Worksheets[iName].Visible = true;
                string iGroup = lv选中.Items[i].Group.Name;
                if (iGroup == "MustGroup")
                {

                        WorkingPaper.Wb.Worksheets[iName].PageSetup.Zoom = false;
                        WorkingPaper.Wb.Worksheets[iName].PageSetup.Orientation =
                            MustDirection[iNo] == "竖向" ? XlPageOrientation.xlPortrait : XlPageOrientation.xlLandscape;
                    WorkingPaper.Wb.Worksheets[iName].PageSetup.PrintArea = MustArea[iNo];
                    HW = MustZoom[iNo].Split(new char[] { '-' });
                    WorkingPaper.Wb.Worksheets[iName].PageSetup.FitToPagesWide = Convert.ToInt16(HW[0]);
                    if (HW[1] == "0")
                    {
                        WorkingPaper.Wb.Worksheets[iName].PageSetup.FitToPagesTall = false;
                    }
                    else
                    {
                        WorkingPaper.Wb.Worksheets[iName].PageSetup.FitToPagesTall = Convert.ToInt16(HW[1]);
                    }
                }
                else if (iGroup == "ChooseGroup")
                {

                        WorkingPaper.Wb.Worksheets[iName].PageSetup.Zoom = false;
                        WorkingPaper.Wb.Worksheets[iName].PageSetup.Orientation =
                            ChooseDirection[iNo] == "竖向" ? XlPageOrientation.xlPortrait : XlPageOrientation.xlLandscape;
                    WorkingPaper.Wb.Worksheets[iName].PageSetup.PrintArea = ChooseArea[iNo];
                    switch (iName)
                    {
                        case "凭证检查":
                            j = WorkingPaper.Wb.Worksheets[iName].Range["M205"].End[XlDirection.xlUp].Row + 1;
                            WorkingPaper.Wb.Worksheets[iName].Rows["1:206"].EntireRow.Hidden = false;
                            WorkingPaper.Wb.Worksheets[iName].Rows[j + ":205"].EntireRow.Hidden = true;
                            break;
                        case "折旧测算":
                            j = WorkingPaper.Wb.Worksheets[iName].Range["A65586"].End[XlDirection.xlUp].Row;
                            WorkingPaper.Wb.Worksheets[iName].PageSetup.PrintArea = "$A$1:$O$" + j;
                            break;
                        case "现金证明":
                                WorkingPaper.Wb.Worksheets[iName].Range["C18"].NumberFormatLocal = "yyyy-mm-dd";
                                break;
                        case "银行调节":
                            WorkingPaper.Wb.Worksheets[iName].Range["A36"].NumberFormatLocal = "yyyy-mm-dd";
                            break;
                            case "应收":
                        case "预付":
                        case "其他应收":
                        case "应付":
                        case "预收":
                        case "其他应付":
                            j = Math.Max(WorkingPaper.Wb.Worksheets[iName].Range["A65586"].End[XlDirection.xlUp].Row,
                                Math.Max(WorkingPaper.Wb.Worksheets[iName].Range["B65586"].End[XlDirection.xlUp].Row,
                                Math.Max(WorkingPaper.Wb.Worksheets[iName].Range["C65586"].End[XlDirection.xlUp].Row,
                                WorkingPaper.Wb.Worksheets[iName].Range["D65586"].End[XlDirection.xlUp].Row)));
                            WorkingPaper.Wb.Worksheets[iName].PageSetup.PrintArea = "$A$1:$F$" + j;
                            break;
                    }

                    HW = ChooseZoom[iNo].Split(new char[] { '-' });
                    WorkingPaper.Wb.Worksheets[iName].PageSetup.FitToPagesWide = Convert.ToInt16(HW[0]);
                    if (HW[1] == "0")
                    {
                        WorkingPaper.Wb.Worksheets[iName].PageSetup.FitToPagesTall = false;
                    }
                    else
                    {
                        WorkingPaper.Wb.Worksheets[iName].PageSetup.FitToPagesTall = Convert.ToInt16(HW[1]);
                    }
                }

                    if (o) Globals.WPToolAddln.Application.PrintCommunication = true;
            }
                label1.Text = "STEP.3  正在进行打印预览...";
                this.Refresh();
                string[] s = lists.ToArray();
                /*for (int i = 0; i < lv待选.Items.Count; i++)
                    {
                        //处理Item 
                        string iName = lv待选.Items[i].SubItems[0].Text;
                        WorkingPaper.Wb.Worksheets[iName].Visible = false;
                    }*/

                WorkingPaper.Wb.Worksheets[s].Select();
            Globals.WPToolAddln.Application.ScreenUpdating = true;
                this.DialogResult = DialogResult.Yes;
                this.Close();
                }
                catch (Exception)
                {

                }
                finally
                {
                    Globals.WPToolAddln.Application.ScreenUpdating = true;
                }

            }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            MessageBox.Show(Globals.WPToolAddln.Application.Version.ToString());
        }

        private void btn识别_Click(object sender, EventArgs e)
        {
            if(MessageBox.Show("是否清除当前筛选的表格，自动选择有数数据？","警告",MessageBoxButtons.YesNo,MessageBoxIcon.Warning)==DialogResult.Yes)
            {
                刷新();
                
            }
        }
    }
}
