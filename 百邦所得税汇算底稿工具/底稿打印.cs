using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace 百邦所得税汇算底稿工具
{
    public partial class 底稿打印 : Form
    {

        string[] MustName, MustArea, MustDirection, MustZoom, ChooseName, ChooseArea, ChooseCondition, 
            ChooseDirection, ChooseZoom, NonName;

        private void btn选中_Click(object sender, EventArgs e)
        {
            MessageBox.Show(lv选中.Items[0].Text);
            MessageBox.Show(lv选中.Items[0].SubItems[0].Text);
            MessageBox.Show(lv选中.Items[0].SubItems[1].Text);
        }

        public 底稿打印()
        {
            InitializeComponent();
            刷新();
        }

        public void 刷新()
        {
            this.lv选中.BeginUpdate();
            this.lv待选.BeginUpdate();
            this.lv选中.Items.Clear();
            this.lv待选.Items.Clear();

            //必选打印
            if (Properties.Settings.Default.MustN != "" ||
                Properties.Settings.Default.MustA != "" ||
                Properties.Settings.Default.MustD != "" ||
                Properties.Settings.Default.MustZ != "")
            {
                MustName = Properties.Settings.Default.MustN.Split(new char[] { '/' });
                MustArea = Properties.Settings.Default.MustA.Split(new char[] { '/' });
                MustDirection = Properties.Settings.Default.MustD.Split(new char[] { '/' });
                MustZoom = Properties.Settings.Default.MustZ.Split(new char[] { '/' });
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
                ChooseName = Properties.Settings.Default.ChooseN.Split(new char[] { '/' });
                ChooseArea = Properties.Settings.Default.ChooseA.Split(new char[] { '/' });
                ChooseCondition = Properties.Settings.Default.ChooseC.Split(new char[] { '/' });
                ChooseDirection = Properties.Settings.Default.ChooseD.Split(new char[] { '/' });
                ChooseZoom = Properties.Settings.Default.ChooseZ.Split(new char[] { '/' });
                string[,] Condition = new string[ChooseName.Length, 1];
                for (int i = 0; i < ChooseName.Length; i++)
                {
                    Condition[i, 0] = ChooseCondition[i];
                }
                WorkingPaper.Wb.Worksheets["首页"].Range["M1:M" + ChooseName.Length].FormulaArray = Condition;
                object[,] V = WorkingPaper.Wb.Worksheets["首页"].Range["M1:M" + ChooseName.Length].Value2;
                    for (int i = 0; i < ChooseName.Length; i++)
                {
                    if(CU.Shuzi(V[i+1,1])==0)
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

                this.lv待选.BeginUpdate();
                NonName = Properties.Settings.Default.NonN.Split(new char[] { '/' });
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

            this.lv选中.EndUpdate();
            this.lv待选.EndUpdate();
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
            //try
            //{
                string[] HW;
                Globals.WPToolAddln.Application.PrintCommunication = false;
                Globals.WPToolAddln.Application.ScreenUpdating = false;

            List<string> lists = new List<string>();
            int n = lv选中.Items.Count;
            for (int i = 0; i < n; i++)
                {
                //处理Item 

                label1.Text = "STEP.2  正在设置打印区域..." + (i+1).ToString() + "/" + n.ToString();
                this.Refresh();
                string iName = lv选中.Items[i].SubItems[0].Text;
                lists.Add(iName);
                int iNo = Convert.ToInt16(lv选中.Items[i].SubItems[2].Text);
                    string iGroup = lv选中.Items[i].Group.Name;
                    if (iGroup == "MustGroup")
                    {
                        WorkingPaper.Wb.Worksheets[iName].PageSetup.PrintArea = MustArea[iNo];
                        WorkingPaper.Wb.Worksheets[iName].PageSetup.Orientation =
                            MustDirection[iNo] == "竖向" ? Orientation.Vertical : Orientation.Horizontal;
                        WorkingPaper.Wb.Worksheets[iName].PageSetup.Zoom = false;
                        HW = MustZoom[iNo].Split(new char[] { '-' });
                        WorkingPaper.Wb.Worksheets[iName].PageSetup.FitToPagesWide = Convert.ToInt16(HW[0]);
                        WorkingPaper.Wb.Worksheets[iName].PageSetup.FitToPagesTall = Convert.ToInt16(HW[1]);
                        WorkingPaper.Wb.Worksheets[iName].PageSetup.BlackAndWhite = true;
                        WorkingPaper.Wb.Worksheets[iName].Visible = true;
                    }
                    else if (iGroup == "ChooseGroup")
                    {
                        WorkingPaper.Wb.Worksheets[iName].PageSetup.PrintArea = ChooseArea[iNo];
                    }
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
                Globals.WPToolAddln.Application.PrintCommunication = true;
                Globals.WPToolAddln.Application.ScreenUpdating = true;
            WorkingPaper.Wb.Worksheets[s].Select();

            this.DialogResult = DialogResult.Yes;
            this.Close();
            /*}
            catch (Exception)
            { }
            finally
            {
                Globals.WPToolAddln.Application.PrintCommunication = true;
                Globals.WPToolAddln.Application.ScreenUpdating = true;
            }*/

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

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
