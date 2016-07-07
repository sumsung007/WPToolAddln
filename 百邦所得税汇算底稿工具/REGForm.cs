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
    public partial class REGForm : Form
    {
        public REGForm()
        {
            InitializeComponent();
            UsernameTextBox.Text = CU.机器码();
                Clipboard.SetDataObject(UsernameTextBox.Text);
                MessageBox.Show("机器码已经复制，或点击机器码旁边的按钮进行复制！");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(UsernameTextBox.Text);
            MessageBox.Show("机器码已经复制！");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(PasswordTextBox.Text!="")
            {
                Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\BaiBang", "Key", PasswordTextBox.Text);
                if(CU.授权检测())
                {
                    MessageBox.Show("注册成功，谢谢！");
                    this.DialogResult = DialogResult.Yes;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("授权码有误，请重新输入！");
                }
            }
            else
            {
                MessageBox.Show("授权码为空，请输入后再试！");
            }
        }
    }
}
