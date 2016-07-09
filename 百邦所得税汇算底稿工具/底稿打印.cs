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
            if(Properties.Settings.Default.MustN!=""||
                Properties.Settings.Default.MustN != ""||
                Properties.Settings.Default.MustN != ""||)
            { 
            string[] MustName = Properties.Settings.Default.MustN.Split(new char[] { '/' });
            string[] MustArea = Properties.Settings.Default.MustA.Split(new char[] { '/' });
            foreach (string N in MustName)
            {
                ListViewItem lvi = new ListViewItem();
                lvi.Group = lv待选.Groups["yyfy"];
                lvi.Text = CU.Zifu(Yeb[yyfy - 1, 1]);
                lvi.SubItems.Add(CU.Zifu(Yeb[yyfy - 1, 2]));
                lvi.SubItems.Add(CU.Shuzi(Yeb[yyfy - 1, 5]).ToString("N"));
                lvi.SubItems.Add(Kmlb[yyfy - 1, 1] == null ? "" : CU.Zifu(Kmlb[yyfy - 1, 1]).
                    Substring(CU.Zifu(Kmlb[yyfy - 1, 1]).IndexOf("-") + 1));
                lvi.SubItems.Add((yyfy - 1).ToString());
                lvi.SubItems.Add("0");
                this.listView1.Items.Add(lvi);
            }
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
