using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace 百邦所得税汇算底稿工具
{
    public partial class 企业信息 : Form
    {
        public 企业信息()
        {
            InitializeComponent();
        }

        private void 企业信息_Load(object sender, EventArgs e)
        {
            try
            {
                if (Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["企业情况"] != null)
                {
                    textBox1.Text = Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["企业情况"].cells[1, 1].value;
                    textBox2.Text = Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["企业情况"].cells[2, 1].value;
                    textBox3.Text = Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["企业情况"].cells[3, 1].value;
                    textBox4.Text = Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["企业情况"].cells[4, 1].value;
                    textBox5.Text = Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["企业情况"].cells[5, 1].value;
                    textBox6.Text = Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["企业情况"].cells[6, 1].value;
                    textBox7.Text = Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["企业情况"].cells[7, 1].value;
                    textBox8.Text = Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["企业情况"].cells[8, 1].value;
                    textBox9.Text = Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["企业情况"].cells[9, 1].value;
                    textBox10.Text = Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["企业情况"].cells[10, 1].value;
                    textBox11.Text = Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["企业情况"].cells[11, 1].value;
                }
            }
            catch
            {
                Worksheet ws = Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets.Add();
                ws.Name = "企业情况";
                ws.Visible = XlSheetVisibility.xlSheetHidden;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["About"].range["B13"].value;
        }
    }
}
