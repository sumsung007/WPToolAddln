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
