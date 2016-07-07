using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace 百邦所得税汇算底稿工具
{
    public partial class Contents : UserControl
    {
        public Contents()
        {
            InitializeComponent();
            显示选项卡("");
            树状图();
        }
        
        void 树状图()
        {
            treeView1.Nodes.Clear();
            string[] Text,Tag;
            TreeNode Tn;

            Tn=treeView1.Nodes.Add("综合类底稿");
            Text =Properties.Resources.Text1.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
            Tag = Properties.Resources.Tag1.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
            for  (int i=0;i<Text.Length;i++)
            {
                TreeNode tn=new TreeNode();
                tn.Tag = Tag[i];
                tn.Text = Text[i];
                Tn.Nodes.Add(tn);
            }
            Tn = treeView1.Nodes.Add("调整类底稿");
            Text = Properties.Resources.Text2.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
            Tag = Properties.Resources.Tag2.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < Text.Length; i++)
            {
                TreeNode tn = new TreeNode();
                tn.Tag = Tag[i];
                tn.Text = Text[i];
                Tn.Nodes.Add(tn);
            }
            Tn = treeView1.Nodes.Add("资产类底稿");
            Text = Properties.Resources.Text5.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
            Tag = Properties.Resources.Tag5.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < Text.Length; i++)
            {
                TreeNode tn = new TreeNode();
                tn.Tag = Tag[i];
                tn.Text = Text[i];
                Tn.Nodes.Add(tn);
            }
            Tn = treeView1.Nodes.Add("负债及权益类底稿");
            Text = Properties.Resources.Text6.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
            Tag = Properties.Resources.Tag6.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < Text.Length; i++)
            {
                TreeNode tn = new TreeNode();
                tn.Tag = Tag[i];
                tn.Text = Text[i];
                Tn.Nodes.Add(tn);
            }
            Tn = treeView1.Nodes.Add("损益类底稿");
            Text = Properties.Resources.Text3.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
            Tag = Properties.Resources.Tag3.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < Text.Length; i++)
            {
                TreeNode tn = new TreeNode();
                tn.Tag = Tag[i];
                tn.Text = Text[i];
                Tn.Nodes.Add(tn);
            }
            Tn = treeView1.Nodes.Add("审核报告");
            Text = Properties.Resources.Text4.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
            Tag = Properties.Resources.Tag4.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < Text.Length; i++)
            {
                TreeNode tn = new TreeNode();
                tn.Tag = Tag[i];
                tn.Text = Text[i];
                Tn.Nodes.Add(tn);
            }
        }

        //

        public void 显示选项卡(string ss)
        {
            switch(ss)
            {
                case "余额表":
                    //tP余额表.Parent = tabControl1;
                    splitContainer1.Panel1Collapsed = false;
                    button6.Visible = true;
                    button7.Visible = true;
                    button8.Visible = true;
                    button9.Visible = false;
                    groupBox1.Text = "余额表";
                    break;
                case "税金申报明细":
                    //tP税费.Parent= tabControl1;
                    splitContainer1.Panel1Collapsed = false;
                    button6.Visible = false;
                    button7.Visible = false;
                    button8.Visible = false;
                    button9.Visible = true;
                    groupBox1.Text = "税费测算";
                    break;
                default:
                    splitContainer1.Panel1Collapsed = true;
                    break;
            }
        }

        private void button2_Click(object sender, EventArgs e)//期间费用
        {
            CU.文件判断();
            if (WorkingPaper.Wb.ActiveSheet.Name == "余额表")
            {
                QJFY qj = new QJFY(WorkingPaper.Wb.ActiveSheet);
                qj.ShowDialog();
            }

        }

        private void button1_Click(object sender, EventArgs e)//报表填写
        {
            WorkingPaper.报表填写();
        }


        private void button3_Click(object sender, EventArgs e)//底稿填写
        {
            if (!CU.文件判断())
                return;
            string str;
            if (WorkingPaper.Wb.Worksheets["基本情况"].Cells[8, 2].Value == "中汇百邦（厦门）税务师事务所有限公司" &&
                WorkingPaper.Wb.Worksheets["档案封面"].Cells[6, 1].Value == "中汇百邦（厦门）税务师事务所有限公司" &&
                WorkingPaper.Wb.Worksheets["基本情况（封面）"].Cells[16, 2].Value == "中汇百邦（厦门）税务师事务所有限公司")
            {
                //自动填表开始
                try
                {
                    WorkingPaper.Wb.Application.ScreenUpdating = false;           //自动填写底稿包含货币资金、往来、费用底稿
                    Worksheet SH = WorkingPaper.Wb.Sheets["余额表"];
                    int N = SH.Cells[SH.UsedRange.Rows.Count + 1, 2].End[XlDirection.xlUp].Row;
                    //SH.Columns[9].Clear();
                    //SH.Range["I2:I" + N].FormulaR1C1 = "= COUNTIF(C[-8], RC[-8] & \"?*\")";
                    object[,] YEB = SH.Range["A2:H" + N.ToString()].Value2;
                    object[,] Kemu = SH.Range["N2:N14"].Value2;
                    int n = 0;

                    if (Kemu[1, 1].ToString() != "")                                     //现金
                    {
                        n = Convert.ToInt16(Kemu[1, 1]) - 1;                  //获取一级科目起始行
                        WorkingPaper.Wb.Worksheets["货币资金"].Cells[6, 4].Value = Math.Round(CU.Shuzi(YEB[n, 7]) - CU.Shuzi(YEB[n, 8]), 2);

                    }

                    if (Kemu[2, 1].ToString() != "")                                        //银行存款
                    {
                        n = Convert.ToInt16(Kemu[2, 1]) - 1;
                        str = CU.Zifu(YEB[n, 1]);
                        int i = 12;
                        double m, sum = Math.Round(CU.Shuzi(YEB[n, 7]) - CU.Shuzi(YEB[n, 8]), 2);
                        do
                        {
                            if ((n == N) || (!CU.Zifu(YEB[n + 1, 1]).Contains(CU.Zifu(YEB[n, 1]))))
                            {
                                m = Math.Round(CU.Shuzi(YEB[n, 7]) - CU.Shuzi(YEB[n, 8]), 2);
                                WorkingPaper.Wb.Worksheets["货币资金"].Cells[i, 4].Value = m;
                                WorkingPaper.Wb.Worksheets["货币资金"].Cells[i, 2].Value = YEB[n, 2].ToString().Trim();
                                sum = sum - m;
                                i = i + 1;
                            }
                            n = n + 1;
                        } while (CU.Zifu(YEB[n, 1]).Contains(str) && (i < 18));
                        if (i == 18)
                            WorkingPaper.Wb.Worksheets["货币资金"].Cells[i, 4].Value = sum;
                    }

                    if (Kemu[3, 1].ToString() != "")                                //应收账款
                    {
                        n = Convert.ToInt16(Kemu[3, 1]) - 1;
                        str = CU.Zifu(YEB[n, 1]);
                        List<string> Mingcheng = new List<string>();
                        List<double> Jiefang = new List<double>();
                        List<double> Daifang = new List<double>();

                        do
                        {
                            if ((n == N) || (!CU.Zifu(YEB[n + 1, 1]).Contains(CU.Zifu(YEB[n, 1]))))
                            {
                                Mingcheng.Add(YEB[n, 2].ToString().Trim());
                                Jiefang.Add(CU.Shuzi(YEB[n, 7]));
                                Daifang.Add(CU.Shuzi(YEB[n, 8]));
                            }
                            n = n + 1;
                        } while (CU.Zifu(YEB[n, 1]).Contains(str));
                        string[,] mingcheng = new string[Mingcheng.Count, 1];
                        double[,] jiefang = new double[Jiefang.Count, 1];
                        double[,] daifang = new double[Daifang.Count, 1];

                        int k = 0;
                        foreach (string s in Mingcheng)
                        {
                            mingcheng[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Jiefang)
                        {
                            jiefang[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Daifang)
                        {
                            daifang[k, 0] = s;
                            k++;
                        }

                        WorkingPaper.Wb.Sheets["应收"].Range["A15:A" + (14 + Mingcheng.Count).ToString()].Value2 = mingcheng;
                        WorkingPaper.Wb.Sheets["应收"].Range["B15:B" + (14 + Jiefang.Count).ToString()].Value2 = jiefang;
                        WorkingPaper.Wb.Sheets["应收"].Range["C15:C" + (14 + Daifang.Count).ToString()].Value2 = daifang;
                    }

                    if (Kemu[4, 1].ToString() != "")                                //预付账款
                    {
                        n = Convert.ToInt16(Kemu[4, 1]) - 1;
                        str = CU.Zifu(YEB[n, 1]);
                        List<string> Mingcheng = new List<string>();
                        List<double> Jiefang = new List<double>();
                        List<double> Daifang = new List<double>();

                        do
                        {
                            if ((n == N) || (!CU.Zifu(YEB[n + 1, 1]).Contains(CU.Zifu(YEB[n, 1]))))
                            {
                                Mingcheng.Add(YEB[n, 2].ToString().Trim());
                                Jiefang.Add(CU.Shuzi(YEB[n, 7]));
                                Daifang.Add(CU.Shuzi(YEB[n, 8]));
                            }
                            n = n + 1;
                        } while (CU.Zifu(YEB[n, 1]).Contains(str));
                        string[,] mingcheng = new string[Mingcheng.Count, 1];
                        double[,] jiefang = new double[Jiefang.Count, 1];
                        double[,] daifang = new double[Daifang.Count, 1];

                        int k = 0;
                        foreach (string s in Mingcheng)
                        {
                            mingcheng[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Jiefang)
                        {
                            jiefang[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Daifang)
                        {
                            daifang[k, 0] = s;
                            k++;
                        }

                        WorkingPaper.Wb.Sheets["预付"].Range["A15:A" + (14 + Mingcheng.Count).ToString()].Value2 = mingcheng;
                        WorkingPaper.Wb.Sheets["预付"].Range["B15:B" + (14 + Jiefang.Count).ToString()].Value2 = jiefang;
                        WorkingPaper.Wb.Sheets["预付"].Range["C15:C" + (14 + Daifang.Count).ToString()].Value2 = daifang;
                    }

                    if (Kemu[5, 1].ToString() != "")                                //其他应收款
                    {
                        n = Convert.ToInt16(Kemu[5, 1]) - 1;
                        str = CU.Zifu(YEB[n, 1]);
                        List<string> Mingcheng = new List<string>();
                        List<double> Jiefang = new List<double>();
                        List<double> Daifang = new List<double>();

                        do
                        {
                            if ((n == N) || (!CU.Zifu(YEB[n + 1, 1]).Contains(CU.Zifu(YEB[n, 1]))))
                            {
                                Mingcheng.Add(YEB[n, 2].ToString().Trim());
                                Jiefang.Add(CU.Shuzi(YEB[n, 7]));
                                Daifang.Add(CU.Shuzi(YEB[n, 8]));
                            }
                            n = n + 1;
                        } while (CU.Zifu(YEB[n, 1]).Contains(str));
                        string[,] mingcheng = new string[Mingcheng.Count, 1];
                        double[,] jiefang = new double[Jiefang.Count, 1];
                        double[,] daifang = new double[Daifang.Count, 1];

                        int k = 0;
                        foreach (string s in Mingcheng)
                        {
                            mingcheng[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Jiefang)
                        {
                            jiefang[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Daifang)
                        {
                            daifang[k, 0] = s;
                            k++;
                        }

                        WorkingPaper.Wb.Sheets["其他应收"].Range["A15:A" + (14 + Mingcheng.Count).ToString()].Value2 = mingcheng;
                        WorkingPaper.Wb.Sheets["其他应收"].Range["B15:B" + (14 + Jiefang.Count).ToString()].Value2 = jiefang;
                        WorkingPaper.Wb.Sheets["其他应收"].Range["C15:C" + (14 + Daifang.Count).ToString()].Value2 = daifang;
                    }

                    if (Kemu[6, 1].ToString() != "")                                //应付账款
                    {
                        n = Convert.ToInt16(Kemu[6, 1]) - 1;
                        str = CU.Zifu(YEB[n, 1]);
                        List<string> Mingcheng = new List<string>();
                        List<double> Jiefang = new List<double>();
                        List<double> Daifang = new List<double>();

                        do
                        {
                            if ((n == N) || (!CU.Zifu(YEB[n + 1, 1]).Contains(CU.Zifu(YEB[n, 1]))))
                            {
                                Mingcheng.Add(YEB[n, 2].ToString().Trim());
                                Jiefang.Add(CU.Shuzi(YEB[n, 7]));
                                Daifang.Add(CU.Shuzi(YEB[n, 8]));
                            }
                            n = n + 1;
                        } while (CU.Zifu(YEB[n, 1]).Contains(str));
                        string[,] mingcheng = new string[Mingcheng.Count, 1];
                        double[,] jiefang = new double[Jiefang.Count, 1];
                        double[,] daifang = new double[Daifang.Count, 1];

                        int k = 0;
                        foreach (string s in Mingcheng)
                        {
                            mingcheng[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Jiefang)
                        {
                            jiefang[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Daifang)
                        {
                            daifang[k, 0] = s;
                            k++;
                        }

                        WorkingPaper.Wb.Sheets["应付"].Range["A13:A" + (12 + Mingcheng.Count).ToString()].Value2 = mingcheng;
                        WorkingPaper.Wb.Sheets["应付"].Range["B13:B" + (12 + Jiefang.Count).ToString()].Value2 = jiefang;
                        WorkingPaper.Wb.Sheets["应付"].Range["C13:C" + (12 + Daifang.Count).ToString()].Value2 = daifang;
                    }

                    if (Kemu[7, 1].ToString() != "")                                //预收账款
                    {
                        n = Convert.ToInt16(Kemu[7, 1]) - 1;
                        str = CU.Zifu(YEB[n, 1]);
                        List<string> Mingcheng = new List<string>();
                        List<double> Jiefang = new List<double>();
                        List<double> Daifang = new List<double>();

                        do
                        {
                            if ((n == N) || (!CU.Zifu(YEB[n + 1, 1]).Contains(CU.Zifu(YEB[n, 1]))))
                            {
                                Mingcheng.Add(YEB[n, 2].ToString().Trim());
                                Jiefang.Add(CU.Shuzi(YEB[n, 7]));
                                Daifang.Add(CU.Shuzi(YEB[n, 8]));
                            }
                            n = n + 1;
                        } while (CU.Zifu(YEB[n, 1]).Contains(str));
                        string[,] mingcheng = new string[Mingcheng.Count, 1];
                        double[,] jiefang = new double[Jiefang.Count, 1];
                        double[,] daifang = new double[Daifang.Count, 1];

                        int k = 0;
                        foreach (string s in Mingcheng)
                        {
                            mingcheng[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Jiefang)
                        {
                            jiefang[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Daifang)
                        {
                            daifang[k, 0] = s;
                            k++;
                        }

                        WorkingPaper.Wb.Sheets["预收"].Range["A13:A" + (12 + Mingcheng.Count).ToString()].Value2 = mingcheng;
                        WorkingPaper.Wb.Sheets["预收"].Range["B13:B" + (12 + Jiefang.Count).ToString()].Value2 = jiefang;
                        WorkingPaper.Wb.Sheets["预收"].Range["C13:C" + (12 + Daifang.Count).ToString()].Value2 = daifang;
                    }

                    if (Kemu[8, 1].ToString() != "")                                //其他应付款
                    {
                        n = Convert.ToInt16(Kemu[8, 1]) - 1;
                        str = CU.Zifu(YEB[n, 1]);
                        List<string> Mingcheng = new List<string>();
                        List<double> Jiefang = new List<double>();
                        List<double> Daifang = new List<double>();

                        do
                        {
                            if ((n == N) || (!CU.Zifu(YEB[n + 1, 1]).Contains(CU.Zifu(YEB[n, 1]))))
                            {
                                Mingcheng.Add(YEB[n, 2].ToString().Trim());
                                Jiefang.Add(CU.Shuzi(YEB[n, 7]));
                                Daifang.Add(CU.Shuzi(YEB[n, 8]));
                            }
                            n = n + 1;
                        } while (CU.Zifu(YEB[n, 1]).Contains(str));
                        string[,] mingcheng = new string[Mingcheng.Count, 1];
                        double[,] jiefang = new double[Jiefang.Count, 1];
                        double[,] daifang = new double[Daifang.Count, 1];

                        int k = 0;
                        foreach (string s in Mingcheng)
                        {
                            mingcheng[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Jiefang)
                        {
                            jiefang[k, 0] = s;
                            k++;
                        }
                        k = 0;
                        foreach (double s in Daifang)
                        {
                            daifang[k, 0] = s;
                            k++;
                        }

                        WorkingPaper.Wb.Sheets["其他应付"].Range["A13:A" + (12 + Mingcheng.Count).ToString()].Value2 = mingcheng;
                        WorkingPaper.Wb.Sheets["其他应付"].Range["B13:B" + (12 + Jiefang.Count).ToString()].Value2 = jiefang;
                        WorkingPaper.Wb.Sheets["其他应付"].Range["C13:C" + (12 + Daifang.Count).ToString()].Value2 = daifang;
                    }


                    if (Kemu[11, 1].ToString() != "")                                //实收公积
                    {
                        n = Convert.ToInt16(Kemu[11, 1]) - 1;
                        str = CU.Zifu(YEB[n, 1]);
                        int i = 7;
                        double m = 0, m1, m2, m3;
                        m1 = Math.Round(CU.Shuzi(YEB[n, 4]) - CU.Shuzi(YEB[n, 3]), 2);//期初数
                        m2 = Math.Round(CU.Shuzi(YEB[n, 6]), 2);//本期增加
                        m3 = Math.Round(CU.Shuzi(YEB[n, 5]), 2);//本期减少
                        do
                        {
                            if ((n == N) || (!CU.Zifu(YEB[n + 1, 1]).Contains(CU.Zifu(YEB[n, 1]))))
                            {
                                m = Math.Round(CU.Shuzi(YEB[n, 7]) - CU.Shuzi(YEB[n, 8]), 2);
                                WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 2].Value = YEB[n, 2].ToString().Trim();
                                WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 3].Value = Math.Round(CU.Shuzi(YEB[n, 4]) - CU.Shuzi(YEB[n, 3]), 2);
                                WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 4].Value = Math.Round(CU.Shuzi(YEB[n, 6]), 2);
                                WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 5].Value = Math.Round(CU.Shuzi(YEB[n, 5]), 2);
                                i = i + 1;
                            }
                            n = n + 1;
                        } while (CU.Zifu(YEB[n, 1]).Contains(str) && i < 12);
                        if (i == 12)
                        {
                            WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 2].Value = "其他股东";
                            WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 3].Value = m1 - (double)WorkingPaper.Wb.Worksheets["实收公积"].Cells[13, 3].Value;
                            WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 4].Value = m2 - (double)WorkingPaper.Wb.Worksheets["实收公积"].Cells[13, 4].Value;
                            WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 5].Value = m3 - (double)WorkingPaper.Wb.Worksheets["实收公积"].Cells[13, 5].Value;
                        }
                    }

                    if (Kemu[13, 1].ToString() != "")                                //盈余公积
                    {
                        n = Convert.ToInt16(Kemu[13, 1]) - 1;
                        str = CU.Zifu(YEB[n, 1]);
                        int i = 19;
                        double m = 0, m1, m2, m3;
                        m1 = Math.Round(CU.Shuzi(YEB[n, 4]) - CU.Shuzi(YEB[n, 3]), 2);//期初数
                        m2 = Math.Round(CU.Shuzi(YEB[n, 6]), 2);//本期增加
                        m3 = Math.Round(CU.Shuzi(YEB[n, 5]), 2);//本期减少
                        do
                        {
                            if ((n == N) || (!CU.Zifu(YEB[n + 1, 1]).Contains(CU.Zifu(YEB[n, 1]))))
                            {
                                m = Math.Round(CU.Shuzi(YEB[n, 7]) - CU.Shuzi(YEB[n, 8]), 2);
                                WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 2].Value = YEB[n, 2].ToString().Trim();
                                WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 3].Value = Math.Round(CU.Shuzi(YEB[n, 4]) - CU.Shuzi(YEB[n, 3]), 2);
                                WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 4].Value = Math.Round(CU.Shuzi(YEB[n, 6]), 2);
                                WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 5].Value = Math.Round(CU.Shuzi(YEB[n, 5]), 2);
                                i = i + 1;
                            }
                            n = n + 1;
                        } while (CU.Zifu(YEB[n, 1]).Contains(str) && i < 23);
                        if (i == 23)
                        {
                            WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 2].Value = "其他盈余公积";
                            WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 3].Value = m1 - (double)WorkingPaper.Wb.Worksheets["实收公积"].Cells[24, 3].Value;
                            WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 4].Value = m2 - (double)WorkingPaper.Wb.Worksheets["实收公积"].Cells[24, 4].Value;
                            WorkingPaper.Wb.Worksheets["实收公积"].Cells[i, 5].Value = m3 - (double)WorkingPaper.Wb.Worksheets["实收公积"].Cells[24, 5].Value;
                        }
                    }
                    WorkingPaper.Wb.Application.ScreenUpdating = true;
                    MessageBox.Show("填写底稿完成");
                }
                catch (Exception ex)
                {
                    Globals.WPToolAddln.Application.ScreenUpdating = true;
                    MessageBox.Show("用户操作出现错误：" + ex.Message);
                }
            }
        }
        

        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)//双击树状图
        {
            if (e.Node.GetNodeCount(false) == 0)
            {
                if (CU.文件判断())
                {

                    WorkingPaper.Wb.Worksheets[e.Node.Tag].Visible = true;
                    WorkingPaper.Wb.Worksheets[e.Node.Tag].Select();
                }
            }
        }

        private void button4_Click(object sender, EventArgs E)
        {
            if (WorkingPaper.OOO)
            {
                if ((WorkingPaper.Wb.Sheets["基本情况"].Cells[8, 2].Value == "中汇百邦（厦门）税务师事务所有限公司") &&
                    (WorkingPaper.Wb.Sheets["档案封面"].Cells[6, 1].Value == "中汇百邦（厦门）税务师事务所有限公司") &&
                    (WorkingPaper.Wb.Sheets["基本情况（封面）"].Cells[16, 2].Value == "中汇百邦（厦门）税务师事务所有限公司"))
                {
                    try
                    {
                        WorkingPaper.Wb.Application.ScreenUpdating = false;
                        Worksheet SH = WorkingPaper.Wb.Sheets["税金申报明细"];
                        SH.Range["A:I"].Replace(SH.Range["R5"].Value, ""); //【税费缴纳测算】表
                        double jj = 0;
                        double[,] c = new double[13, 1], e = new double[13, 1], m = new double[7, 1], k = new double[6, 1];
                        int n = SH.Cells[SH.UsedRange.Rows.Count + 1, 1].End[XlDirection.xlUp].Row;
                        object[,] Shuifei = SH.Range["A2:I" + n.ToString()].Value2;
                        for (int i = 1; i <= n - 1; i++)
                        {
                            if (Shuifei[i, 2] != null)
                            {
                                string str = CU.Zifu(Shuifei[i, 2]);
                                if (str.Contains("印花税"))
                                {
                                    switch (str)
                                    {
                                        case "印花税-购销合同":
                                            c[0, 0] = c[0, 0] + CU.Shuzi(Shuifei[i, 4]);
                                            e[0, 0] = e[0, 0] + CU.Shuzi(Shuifei[i, 8]);
                                            break;
                                        case "印花税-建筑安装工程承包合同":
                                            c[1, 0] = c[1, 0] + CU.Shuzi(Shuifei[i, 4]);
                                            e[1, 0] = e[1, 0] + CU.Shuzi(Shuifei[i, 8]);
                                            break;
                                        case "印花税-技术合同":
                                            c[2, 0] = c[2, 0] + CU.Shuzi(Shuifei[i, 4]);
                                            e[2, 0] = e[2, 0] + CU.Shuzi(Shuifei[i, 8]);
                                            break;
                                        case "印花税-财产租赁合同":
                                            c[3, 0] = c[3, 0] + CU.Shuzi(Shuifei[i, 4]);
                                            e[3, 0] = e[3, 0] + CU.Shuzi(Shuifei[i, 8]);
                                            break;
                                        case "印花税-仓储保管合同":
                                            c[4, 0] = c[4, 0] + CU.Shuzi(Shuifei[i, 4]);
                                            e[4, 0] = e[4, 0] + CU.Shuzi(Shuifei[i, 8]);
                                            break;
                                        case "印花税-财产保险合同":
                                            c[5, 0] = c[5, 0] + CU.Shuzi(Shuifei[i, 4]);
                                            e[5, 0] = e[5, 0] + CU.Shuzi(Shuifei[i, 8]);
                                            break;
                                        case "印花税-货物运输合同":
                                            c[6, 0] = c[6, 0] + CU.Shuzi(Shuifei[i, 4]);
                                            e[6, 0] = e[6, 0] + CU.Shuzi(Shuifei[i, 8]);
                                            break;
                                        case "印花税-加工承揽合同":
                                            c[7, 0] = c[7, 0] + CU.Shuzi(Shuifei[i, 4]);
                                            e[7, 0] = e[7, 0] + CU.Shuzi(Shuifei[i, 8]);
                                            break;
                                        case "印花税-建设工程勘察设计合同":
                                            c[8, 0] = c[8, 0] + CU.Shuzi(Shuifei[i, 4]);
                                            e[8, 0] = e[8, 0] + CU.Shuzi(Shuifei[i, 8]);
                                            break;
                                        case "印花税-产权转移书据":
                                            c[9, 0] = c[9, 0] + CU.Shuzi(Shuifei[i, 4]);
                                            e[9, 0] = e[9, 0] + CU.Shuzi(Shuifei[i, 8]);
                                            break;
                                        case "印花税-借款合同":
                                            c[10, 0] = c[10, 0] + CU.Shuzi(Shuifei[i, 4]);
                                            e[10, 0] = e[10, 0] + CU.Shuzi(Shuifei[i, 8]);
                                            break;
                                        case "印花税-其他营业帐簿":
                                        case "权利、许可证照":
                                        case "印花税-经财政部确定的其他凭证":
                                            c[11, 0] = c[11, 0] + CU.Shuzi(Shuifei[i, 4]);
                                            e[11, 0] = e[11, 0] + CU.Shuzi(Shuifei[i, 8]);
                                            break;
                                        case "印花税-资金帐簿":
                                            c[12, 0] = c[12, 0] + CU.Shuzi(Shuifei[i, 4]);
                                            e[12, 0] = e[12, 0] + CU.Shuzi(Shuifei[i, 8]);
                                            break;
                                        default:
                                            break;
                                    }
                                }
                                else
                                {
                                    if (str.Contains("消费税"))
                                    {
                                        m[0, 0] = m[0, 0] + CU.Shuzi(Shuifei[i, 8]);
                                    }
                                    else
                                    {
                                        if (str.Contains("营业税"))
                                        {
                                            m[1, 0] = m[1, 0] + CU.Shuzi(Shuifei[i, 8]);
                                        }
                                        else
                                        {
                                            if (str.Contains("城建税"))
                                            {
                                                m[2, 0] = m[2, 0] + CU.Shuzi(Shuifei[i, 8]);
                                            }
                                            else
                                            {
                                                if (str.Contains("教育费附加"))
                                                {
                                                    m[3, 0] = m[3, 0] + CU.Shuzi(Shuifei[i, 8]);
                                                }
                                                else
                                                {
                                                    if (str.Contains("地方教育附加"))
                                                    {
                                                        m[4, 0] = m[4, 0] + CU.Shuzi(Shuifei[i, 8]);
                                                    }
                                                    else
                                                    {
                                                        if (str.Contains("资源税"))
                                                        {
                                                            m[5, 0] = m[5, 0] + CU.Shuzi(Shuifei[i, 8]);
                                                        }
                                                        else
                                                        {
                                                            if (str.Contains("土地增值税"))
                                                            {
                                                                m[6, 0] = m[6, 0] + CU.Shuzi(Shuifei[i, 8]);
                                                            }
                                                            else
                                                            {
                                                                if (str.Contains("房产税"))
                                                                {
                                                                    jj = jj + CU.Shuzi(Shuifei[i, 8]);
                                                                }
                                                                else
                                                                {
                                                                    if (str.Contains("车船税"))
                                                                    {
                                                                        k[0, 0] = k[0, 0] + CU.Shuzi(Shuifei[i, 8]);
                                                                    }
                                                                    else
                                                                    {
                                                                        if (str.Contains("土地使用税"))
                                                                        {
                                                                            k[1, 0] = k[1, 0] + CU.Shuzi(Shuifei[i, 8]);
                                                                        }
                                                                        else
                                                                        {
                                                                            if (str.Contains("契税"))
                                                                            {
                                                                                k[2, 0] = k[2, 0] + CU.Shuzi(Shuifei[i, 8]);
                                                                            }
                                                                            else
                                                                            {
                                                                                if (str.Contains("残疾人就业"))
                                                                                {
                                                                                    k[3, 0] = k[3, 0] + CU.Shuzi(Shuifei[i, 8]);
                                                                                }
                                                                                else
                                                                                {
                                                                                    if (str.Contains("文化事业建设费"))
                                                                                    {
                                                                                        k[4, 0] = k[4, 0] + CU.Shuzi(Shuifei[i, 8]);
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        if (str.Contains("个人所得税"))
                                                                                        {
                                                                                            k[5, 0] = k[5, 0] + CU.Shuzi(Shuifei[i, 8]);
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        c[11, 0] = e[11, 0] / 5;

                        WorkingPaper.Wb.Sheets["税费缴纳测算"].Range["C37:C49"].Value2 = c;
                        WorkingPaper.Wb.Sheets["税费缴纳测算"].Range["E37:E49"].Value2 = e;
                        WorkingPaper.Wb.Sheets["税费缴纳测算"].Cells[17, 5].Value = e[0, 0] + e[1, 0] + e[2, 0] + e[3, 0] +
                            e[4, 0] + e[5, 0] + e[6, 0] + e[7, 0] + e[8, 0] + e[9, 0] + e[10, 0] + e[11, 0];

                        WorkingPaper.Wb.Sheets["税费缴纳测算"].Range["E8:E14"].Value2 = m;
                        WorkingPaper.Wb.Sheets["税费缴纳测算"].Cells[16, 5].Value = jj;
                        WorkingPaper.Wb.Sheets["税费缴纳测算"].Range["E18:E23"].Value2 = k;
                        WorkingPaper.Wb.Application.ScreenUpdating = true;
                        MessageBox.Show("税费填写完成");
                    }
                    catch (Exception ex)
                    {
                        Globals.WPToolAddln.Application.ScreenUpdating = true;
                        MessageBox.Show("用户操作出现错误：" + ex.Message);
                    }
                }
            }
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}
