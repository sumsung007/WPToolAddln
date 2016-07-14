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
                string[] MustName = Properties.Settings.Default.MustN.Split(new char[] { '/' });
                string[] MustArea = Properties.Settings.Default.MustA.Split(new char[] { '/' });
                string[] MustDirection = Properties.Settings.Default.MustD.Split(new char[] { '/' });
                string[] MustZoom = Properties.Settings.Default.MustZ.Split(new char[] { '/' });
                for (int i = 0; i < MustName.Length; i++)
                {
                    ListViewItem lvi = new ListViewItem();
                    lvi.Group = lv选中.Groups["MustGroup"];
                    lvi.Text = MustName[i];
                    lvi.SubItems.Add("必选");
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
                string[] ChooseName = Properties.Settings.Default.ChooseN.Split(new char[] { '/' });
                string[] ChooseArea = Properties.Settings.Default.ChooseA.Split(new char[] { '/' });
                string[] ChooseCondition = Properties.Settings.Default.ChooseC.Split(new char[] { '/' });
                string[] ChooseDirection = Properties.Settings.Default.ChooseD.Split(new char[] { '/' });
                string[] ChooseZoom = Properties.Settings.Default.ChooseZ.Split(new char[] { '/' });
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
                        this.lv待选.Items.Add(lvi);

                    }
                    else
                    {
                        ListViewItem lvi = new ListViewItem();
                        lvi.Group = lv选中.Groups["ChooseGroup"];
                        lvi.Text = ChooseName[i];
                        lvi.SubItems.Add("有数");
                        this.lv选中.Items.Add(lvi);
                    }
                }
            }

            //不用打印
            if (Properties.Settings.Default.NonN != "")
            {

                this.lv待选.BeginUpdate();
                string[] NonName = Properties.Settings.Default.NonN.Split(new char[] { '/' });
                for (int i = 0; i < NonName.Length; i++)
                {
                    ListViewItem lvi = new ListViewItem();
                    lvi.Group = lv待选.Groups["NonGroup"];
                    lvi.Text = NonName[i];
                    lvi.SubItems.Add("无需");
                    this.lv待选.Items.Add(lvi);
                }
            }

            this.lv选中.EndUpdate();
            this.lv待选.EndUpdate();
        }

        private void btn打印_Click(object sender, EventArgs e)
        {
            Globals.WPToolAddln.Application.ActiveWorkbook.ActiveSheet.Range["A1:D10"].Copy();
            Image img;
            if (System.Windows.Forms.Clipboard.ContainsImage())
            {
                img = System.Windows.Forms.Clipboard.GetImage();
                pictureBox1.Image = img;

            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
