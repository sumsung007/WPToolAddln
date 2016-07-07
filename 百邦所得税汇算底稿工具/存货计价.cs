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
    public partial class 存货计价 : Form
    {
        public 存货计价()
        {
            InitializeComponent();
            if(WorkingPaper.Wb.ActiveSheet.Range["B15:F15"].Value!=null)
            {
                string str = WorkingPaper.Wb.ActiveSheet.Range["B15"].Value.ToString();
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemChecked(i, str.Contains(checkedListBox1.Items[i].ToString()));
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string str = "";
            for(int i=0;i<checkedListBox1.Items.Count;i++)
            {
                if(checkedListBox1.GetItemChecked(i))
                {
                    str = str + ","+ checkedListBox1.Items[i].ToString();
                }
            }
            if(str=="")
            {
                WorkingPaper.Wb.ActiveSheet.Range["B15:F15"].Value=str;
            }
            else
            {
                WorkingPaper.Wb.ActiveSheet.Range["B15:F15"].Value = str.Substring(1,str.Length-1);
            }
            this.Close();
        }
    }
}
