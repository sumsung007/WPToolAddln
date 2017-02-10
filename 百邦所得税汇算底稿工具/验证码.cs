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
    public partial class 验证码 : Form
    {
        public string pictext;
        public 验证码(Image img,string cp)
        {
            InitializeComponent();
            pictureBox1.Image = img;
            this.Text = cp;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            pictext = textBox1.Text;
            this.Close();
        }
    }
}
