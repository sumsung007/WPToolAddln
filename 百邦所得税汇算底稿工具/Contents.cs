using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using DataTable = System.Data.DataTable;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;
using Newtonsoft.Json;

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
                    splitContainer3.Panel1Collapsed = false;
                    splitContainer3.Panel2Collapsed = true;
                    groupBox1.Text = "余额表";
                    break;
                case "税金申报明细":
                    //tP税费.Parent= tabControl1;
                    splitContainer1.Panel1Collapsed = false;
                    splitContainer3.Panel2Collapsed = false;
                    splitContainer3.Panel1Collapsed = true;
                    splitContainer4.Panel1Collapsed = false;
                    splitContainer4.Panel2Collapsed = true;
                    groupBox1.Text = "税费测算";
                    break;
                case "基本情况":
                    //tP税费.Parent= tabControl1;
                    splitContainer1.Panel1Collapsed = false;
                    splitContainer3.Panel2Collapsed = false;
                    splitContainer3.Panel1Collapsed = true;
                    splitContainer4.Panel2Collapsed = false;
                    splitContainer4.Panel1Collapsed = true;
                    groupBox1.Text = "基本情况";
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
            if ((WorkingPaper.Wb.Worksheets["基本情况"].Cells[8, 2].Value == "中汇百邦（厦门）税务师事务所有限公司" &&
                WorkingPaper.Wb.Worksheets["档案封面"].Cells[6, 1].Value == "中汇百邦（厦门）税务师事务所有限公司" &&
                WorkingPaper.Wb.Worksheets["基本情况（封面）"].Cells[16, 2].Value == "中汇百邦（厦门）税务师事务所有限公司")||
                (WorkingPaper.Wb.Worksheets["基本情况"].Cells[8, 2].Value == "厦门明正税务师事务所有限公司" &&
                WorkingPaper.Wb.Worksheets["档案封面"].Cells[6, 1].Value == "厦门明正税务师事务所有限公司" &&
                WorkingPaper.Wb.Worksheets["基本情况（封面）"].Cells[16, 2].Value == "厦门明正税务师事务所有限公司"))
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
                if (((WorkingPaper.Wb.Sheets["基本情况"].Cells[8, 2].Value == "中汇百邦（厦门）税务师事务所有限公司") &&
                    (WorkingPaper.Wb.Sheets["档案封面"].Cells[6, 1].Value == "中汇百邦（厦门）税务师事务所有限公司") &&
                    (WorkingPaper.Wb.Sheets["基本情况（封面）"].Cells[16, 2].Value == "中汇百邦（厦门）税务师事务所有限公司"))||
                    ((WorkingPaper.Wb.Sheets["基本情况"].Cells[8, 2].Value == "厦门明正税务师事务所有限公司") &&
                    (WorkingPaper.Wb.Sheets["档案封面"].Cells[6, 1].Value == "厦门明正税务师事务所有限公司") &&
                    (WorkingPaper.Wb.Sheets["基本情况（封面）"].Cells[16, 2].Value == "厦门明正税务师事务所有限公司")
                    ))
                {
                    try
                    {
                        WorkingPaper.Wb.Application.ScreenUpdating = false;
                        Worksheet SH = WorkingPaper.Wb.Sheets["税金申报明细"];
                        //SH.Range["A:I"].Replace(SH.Range["R5"].Value, ""); //【税费缴纳测算】表
                        double jj = 0;
                        double[,] c = new double[13, 1], e = new double[13, 1], m = new double[7, 1], k = new double[6, 1];
                        int n = SH.Cells[SH.UsedRange.Rows.Count + 1, 1].End[XlDirection.xlUp].Row;
                        object[,] Shuifei = SH.Range["A2:L" + n.ToString()].Value2;
                        string year = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["B4"].Value2);
                        for (int i = 1; i <= n - 1; i++)
                        {
                            if (Shuifei[i, 2] != null && CU.Zifu(Shuifei[i,5]).Substring(0,4)==year)
                            {
                                string 征收项目 = CU.Zifu(Shuifei[i, 2]);
                                if (征收项目.Contains("印花税"))
                                {
                                    string 征收品目 = CU.Zifu(Shuifei[i, 3]);
                                    switch (征收品目)
                                    {
                                        case "购销合同":
                                            c[0, 0] = c[0, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[0, 0] = e[0, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "建筑安装工程承包合同":
                                            c[1, 0] = c[1, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[1, 0] = e[1, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "技术合同":
                                            c[2, 0] = c[2, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[2, 0] = e[2, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "财产租赁合同":
                                            c[3, 0] = c[3, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[3, 0] = e[3, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "仓储保管合同":
                                            c[4, 0] = c[4, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[4, 0] = e[4, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "财产保险合同":
                                            c[5, 0] = c[5, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[5, 0] = e[5, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "货物运输合同(按运输费用万分之五贴花)":
                                            c[6, 0] = c[6, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[6, 0] = e[6, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "加工承揽合同":
                                            c[7, 0] = c[7, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[7, 0] = e[7, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "建设工程勘察设计合同":
                                            c[8, 0] = c[8, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[8, 0] = e[8, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "产权转移书据":
                                            c[9, 0] = c[9, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[9, 0] = e[9, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "借款合同":
                                            c[10, 0] = c[10, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[10, 0] = e[10, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "其他营业账簿":
                                        case "权利、许可证照":
                                        case "其他凭证":
                                            c[11, 0] = c[11, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[11, 0] = e[11, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        case "资金账簿":
                                            c[12, 0] = c[12, 0] + CU.Shuzi(Shuifei[i, 6]);
                                            e[12, 0] = e[12, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            break;
                                        default:
                                            break;
                                    }
                                }
                                else
                                {
                                    if (征收项目.Contains("消费税"))
                                    {
                                        m[0, 0] = m[0, 0] + CU.Shuzi(Shuifei[i, 11]);
                                    }
                                    else
                                    {
                                        if (征收项目.Contains("营业税"))
                                        {
                                            m[1, 0] = m[1, 0] + CU.Shuzi(Shuifei[i, 11]);
                                        }
                                        else
                                        {
                                            if (征收项目.Contains("城市维护建设税"))
                                            {
                                                m[2, 0] = m[2, 0] + CU.Shuzi(Shuifei[i, 11]);
                                            }
                                            else
                                            {
                                                if (征收项目.Contains("教育费附加"))
                                                {
                                                    m[3, 0] = m[3, 0] + CU.Shuzi(Shuifei[i, 11]);
                                                }
                                                else
                                                {
                                                    if (征收项目.Contains("地方教育附加"))
                                                    {
                                                        m[4, 0] = m[4, 0] + CU.Shuzi(Shuifei[i, 11]);
                                                    }
                                                    else
                                                    {
                                                        if (征收项目.Contains("资源税"))
                                                        {
                                                            m[5, 0] = m[5, 0] + CU.Shuzi(Shuifei[i, 11]);
                                                        }
                                                        else
                                                        {
                                                            if (征收项目.Contains("土地增值税"))
                                                            {
                                                                m[6, 0] = m[6, 0] + CU.Shuzi(Shuifei[i, 11]);
                                                            }
                                                            else
                                                            {
                                                                if (征收项目.Contains("房产税"))
                                                                {
                                                                    jj = jj + CU.Shuzi(Shuifei[i, 11]);
                                                                }
                                                                else
                                                                {
                                                                    if (征收项目.Contains("车船税"))
                                                                    {
                                                                        k[0, 0] = k[0, 0] + CU.Shuzi(Shuifei[i, 11]);
                                                                    }
                                                                    else
                                                                    {
                                                                        if (征收项目.Contains("城镇土地使用税"))
                                                                        {
                                                                            k[1, 0] = k[1, 0] + CU.Shuzi(Shuifei[i, 11]);
                                                                        }
                                                                        else
                                                                        {
                                                                            if (征收项目.Contains("契税"))
                                                                            {
                                                                                k[2, 0] = k[2, 0] + CU.Shuzi(Shuifei[i, 11]);
                                                                            }
                                                                            else
                                                                            {
                                                                                if (征收项目.Contains("残疾人就业保障金"))
                                                                                {
                                                                                    k[3, 0] = k[3, 0] + CU.Shuzi(Shuifei[i, 11]);
                                                                                }
                                                                                else
                                                                                {
                                                                                    if (征收项目.Contains("文化事业建设费"))
                                                                                    {
                                                                                        k[4, 0] = k[4, 0] + CU.Shuzi(Shuifei[i, 11]);
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        if (征收项目.Contains("个人所得税"))
                                                                                        {
                                                                                            k[5, 0] = k[5, 0] + CU.Shuzi(Shuifei[i, 11]);
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

        #region 基本情况
        private void btn基础信息_Click(object sender, EventArgs e)
        {
            if (!CU.文件判断())
                return;
            string name1, pass1;
            name1 = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["B49"].Value2);
            pass1 = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["D49"].Value2);
            if (name1 == "" || pass1 == "")
                MessageBox.Show("国税用户名密码未填写，请填写[基本情况].[B49,D49]后重试！");
            else if (国税信息(name1,base64(pass1)))
                MessageBox.Show("国税信息抓取成功！");
            else
                MessageBox.Show("国税抓取失败！");
            name1 = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["B50"].Value2);
            pass1 = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["D50"].Value2);
            if (name1 == "" || pass1 == "")
                MessageBox.Show("地税用户名密码未填写，请填写[基本情况].[B50,D50]后重试！");
            else if (地税信息(name1,pass1))
                MessageBox.Show("地税信息抓取成功！");
            else
                MessageBox.Show("地税抓取失败！");
        }
        #endregion

        #region 申报数据获取

        private void button1_Click_1(object sender, EventArgs e)
        {
            string strText, scookie, strName, strPass;
            if (!CU.文件判断())
                return;
            strName = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["B50"].Value2);
            strPass = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["D50"].Value2);
            if (strName == "" || strPass == "")
            {
                MessageBox.Show("地税用户名和密码未填写，请填写[基本情况].[B50,D50]后重试！");
                return;
            }
            
            HttpHelper http = new HttpHelper();
            HttpItem item = new HttpItem()
            {
                URL = "https://www.xm-l-tax.gov.cn/", //URL     必需项
                IsToLower = false, //得到的HTML代码是否转成小写     可选项默认转小写
                UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
            };
            HttpResult result = http.GetHtml(item);
            scookie = result.Cookie;

            item = new HttpItem
            {
                URL = "https://www.xm-l-tax.gov.cn/common/checkcode.do?rand=" + System.DateTime.Now.ToString("ddd MMM dd hh:mm:ss \"CST\" yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo),
                //URL     必需项
                Referer = "https://www.xm-l-tax.gov.cn/", //来源URL     可选项
                IsToLower = false, //得到的HTML代码是否转成小写     可选项默认转小写
                Cookie = scookie,
                ResultType = ResultType.Byte, //返回数据类型，是Byte还是String
                UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
            };
            result = http.GetHtml(item);

            //把得到的Byte转成图片
            Image img = byteArrayToImage(result.ResultByte);
            验证码 pic = new 验证码(img,"地税验证码");
            pic.StartPosition = FormStartPosition.CenterParent;
            pic.ShowDialog();
            strText = pic.pictext;
            //strPass = md5(strPass);
            strPass = base64(strPass);

            item = new HttpItem
            {
                URL = "https://www.xm-l-tax.gov.cn/login/checkLogin.do", //URL     必需项
                Method = "post", //URL     可选项 默认为Get
                Referer = "https://www.xm-l-tax.gov.cn/", //来源URL     可选项
                Cookie = scookie,
                IsToLower = false, //得到的HTML代码是否转成小写     可选项默认转小写
                ContentType = "application/x-www-form-urlencoded; charset=UTF-8", //返回类型    可选项有默认值
                UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                Postdata =
                    "loginId=" + strName + "&userPassword=" + strPass + "&checkCode=" + strText, //Post数据
            };
            item.Header.Add("x-requested-with", "XMLHttpRequest");

            result = http.GetHtml(item);

            if (result.Html.ToString().IndexOf("登录成功") > 0)
            {
                item = new HttpItem
                {
                    URL = "https://www.xm-l-tax.gov.cn/login/gdslhrz.do", //URL     必需项
                    Method = "post", //URL     可选项 默认为Get
                    Referer = "https://www.xm-l-tax.gov.cn/", //来源URL     可选项
                    Cookie = scookie,
                    IsToLower = false, //得到的HTML代码是否转成小写     可选项默认转小写
                    ContentType = "application/x-www-form-urlencoded; charset=UTF-8", //返回类型    可选项有默认值
                    UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                    Postdata =
                        "loginId=" + strName + "&userPassword=" + strPass + "&checkCode=" + strText, //Post数据
                };
                item.Header.Add("x-requested-with", "XMLHttpRequest");
                result = http.GetHtml(item);

                item = new HttpItem
                {
                    URL = "https://www.xm-l-tax.gov.cn/login/opener.do", //URL     必需项
                    Referer = "https://www.xm-l-tax.gov.cn/", //来源URL     可选项
                    Cookie = scookie,
                    UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                };
                result = http.GetHtml(item);
                item = new HttpItem
                {
                    URL = "https://www.xm-l-tax.gov.cn/nsfwHome/index.do", //URL     必需项
                    Cookie = scookie,
                    UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                };
                result = http.GetHtml(item);
                item = new HttpItem
                {
                    URL = "https://www.xm-l-tax.gov.cn/xxtx/index.do", //URL     必需项
                    Cookie = scookie,
                    UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                };
                result = http.GetHtml(item);
                item = new HttpItem
                {
                    URL = "https://www.xm-l-tax.gov.cn/nsfwHome/index.do?menuid=zhcx", //URL     必需项
                    Referer = "https://www.xm-l-tax.gov.cn/xxtx/index.do", //来源URL     可选项
                    Cookie = scookie,
                    UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                };
                result = http.GetHtml(item);
                item = new HttpItem
                {
                    URL = "https://www.xm-l-tax.gov.cn/dzsb/query/qnsssbrk_index.do", //URL     必需项
                    Referer = "https://www.xm-l-tax.gov.cn/nsfwHome/index.do?menuid=zhcx", //来源URL     可选项
                    Cookie = scookie,
                    UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                };
                result = http.GetHtml(item);
                string year = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["B4"].Value2);
                item = new HttpItem
                {
                    URL = "https://www.xm-l-tax.gov.cn/dzsb/query/qnsssbrk_query.do", //URL     必需项
                    Method = "post", //URL     可选项 默认为Get
                    Referer = "https://www.xm-l-tax.gov.cn/dzsb/query/qnsssbrk_index.do", //来源URL     可选项
                    Cookie = scookie,
                    IsToLower = false, //得到的HTML代码是否转成小写     可选项默认转小写
                    ContentType = "application/json;charset=UTF-8", //返回类型    可选项有默认值
                    UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                    Postdata =
                        "{\"zfbz\":\"N\",\"sbrq_year\":\"" + year + "\",\"sbrq_month\":\"\",\"queryType\":\"byyear\"}",
                    //Post数据
                };
                item.Header.Add("x-requested-with", "XMLHttpRequest");

                result = http.GetHtml(item);

                string html = result.Html;
                CJson cj = JsonConvert.DeserializeObject<CJson>(html);
                lsWssbjl[] dates1 = cj.lsWssbjl;


                item = new HttpItem
                {
                    URL = "https://www.xm-l-tax.gov.cn/dzsb/query/qnsssbrk_query.do", //URL     必需项
                    Method = "post", //URL     可选项 默认为Get
                    Referer = "https://www.xm-l-tax.gov.cn/dzsb/query/qnsssbrk_index.do", //来源URL     可选项
                    Cookie = scookie,
                    IsToLower = false, //得到的HTML代码是否转成小写     可选项默认转小写
                    ContentType = "application/json;charset=UTF-8", //返回类型    可选项有默认值
                    UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                    Postdata =
                        "{\"zfbz\":\"N\",\"sbrq_year\":\"" + (Convert.ToInt16(year) + 1).ToString() +
                        "\",\"sbrq_month\":\"1\",\"queryType\":\"bymonth\"}", //Post数据
                };
                item.Header.Add("x-requested-with", "XMLHttpRequest");

                result = http.GetHtml(item);
                html = result.Html;
                cj = JsonConvert.DeserializeObject<CJson>(html);
                lsWssbjl[] dates2 = cj.lsWssbjl;
                string[,] dr = new string[dates1.Length + dates2.Length, 12];
                for (int i = 0; i < dates1.Length; i++)
                {
                    dr[i, 0] = dates1[i].nssbrq;
                    dr[i, 1] = dates1[i].zsxm_dm;
                    dr[i, 2] = dates1[i].zspm_dm;
                    dr[i, 3] = dates1[i].zszm_dm;
                    dr[i, 4] = dates1[i].skssqq + "到" + dates1[i].skssqz;
                    dr[i, 5] = dates1[i].jsyj;
                    dr[i, 6] = dates1[i].sl_1;
                    dr[i, 7] = dates1[i].ynse;
                    dr[i, 8] = dates1[i].yjse;
                    dr[i, 9] = dates1[i].jmse;
                    dr[i, 10] = dates1[i].ybtse;
                    dr[i, 11] = dates1[i].rkrq;
                }
                int j = dates1.Length;
                for (int i = 0; i < dates2.Length; i++)
                {
                    dr[j + i, 0] = dates2[i].nssbrq;
                    dr[j + i, 1] = dates2[i].zsxm_dm;
                    dr[j + i, 2] = dates2[i].zspm_dm;
                    dr[j + i, 3] = dates2[i].zszm_dm;
                    dr[j + i, 4] = dates2[i].skssqq + "到" + dates2[i].skssqz;
                    dr[j + i, 5] = dates2[i].jsyj;
                    dr[j + i, 6] = dates2[i].sl_1;
                    dr[j + i, 7] = dates2[i].ynse;
                    dr[j + i, 8] = dates2[i].yjse;
                    dr[j + i, 9] = dates2[i].jmse;
                    dr[j + i, 10] = dates2[i].ybtse;
                    dr[j + i, 11] = dates2[i].rkrq;
                }

                WorkingPaper.Wb.Application.ScreenUpdating = false;
                int lRow = WorkingPaper.Wb.Worksheets["税金申报明细"].Range["A100086"].End[XlDirection.xlUp].Row;
                if (lRow == 1) lRow = 2;
                WorkingPaper.Wb.Worksheets["税金申报明细"].Range["A2:N" + lRow.ToString()].Clear();
                Range rng = WorkingPaper.Wb.Worksheets["税金申报明细"].Range["A2"].Resize[dr.GetLength(0), dr.GetLength(1)];
                rng.Value2 = dr;
                rng = WorkingPaper.Wb.Worksheets["税金申报明细"].Range["M2"].Resize[dr.GetLength(0), 1];
                rng.FormulaR1C1 = "=VLOOKUP(RC[-11],首页!C[-6]:C[-5],2,0)";
                rng = WorkingPaper.Wb.Worksheets["税金申报明细"].Range["N2"].Resize[dr.GetLength(0), 1];
                rng.FormulaR1C1 = "=VLOOKUP(RC[-11],首页!C[-5]:C[-4],2,0)";
                object[,] arr = WorkingPaper.Wb.Worksheets["税金申报明细"].Range["M2"].Resize[dr.GetLength(0), 2].Value2;
                WorkingPaper.Wb.Worksheets["税金申报明细"].Range["B2"].Resize[dr.GetLength(0), 2].Value2 = arr;
                WorkingPaper.Wb.Worksheets["税金申报明细"].Range["M2"].Resize[dr.GetLength(0), 2].Clear();

                WorkingPaper.Wb.Application.ScreenUpdating = true;
                MessageBox.Show("地税申报数据拉取成功");
            }
            else
            {
                MessageBox.Show(result.Html);
            }
        }
        #endregion

        #region 网抓函数
        private Boolean 地税信息(string strName,string strPass)
        {
            string strText, scookie;
            HttpHelper http = new HttpHelper();
            HttpItem item = new HttpItem()
            {
                URL = "https://www.xm-l-tax.gov.cn/", //URL     必需项
                IsToLower = false, //得到的HTML代码是否转成小写     可选项默认转小写
                UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
            };
            HttpResult result = http.GetHtml(item);
            scookie = result.Cookie;

            item = new HttpItem
            {
                URL = "https://www.xm-l-tax.gov.cn/common/checkcode.do?rand=" + System.DateTime.Now.ToString("ddd MMM dd hh:mm:ss \"CST\" yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo),
                //URL     必需项
                Referer = "https://www.xm-l-tax.gov.cn/", //来源URL     可选项
                IsToLower = false, //得到的HTML代码是否转成小写     可选项默认转小写
                Cookie = scookie,
                ResultType = ResultType.Byte, //返回数据类型，是Byte还是String
                UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
            };
            result = http.GetHtml(item);

            //把得到的Byte转成图片
            Image img = byteArrayToImage(result.ResultByte);
            验证码 pic = new 验证码(img, "地税验证码");
            pic.StartPosition = FormStartPosition.CenterParent;
            pic.ShowDialog();
            strText = pic.pictext;
            strPass = base64(strPass);

            item = new HttpItem
            {
                URL = "https://www.xm-l-tax.gov.cn/login/checkLogin.do", //URL     必需项
                Method = "post", //URL     可选项 默认为Get
                Referer = "https://www.xm-l-tax.gov.cn/", //来源URL     可选项
                Cookie = scookie,
                IsToLower = false, //得到的HTML代码是否转成小写     可选项默认转小写
                ContentType = "application/x-www-form-urlencoded; charset=UTF-8", //返回类型    可选项有默认值
                UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                Postdata =
                    "loginId=" + strName + "&userPassword=" + strPass + "&checkCode=" + strText, //Post数据
            };
            item.Header.Add("x-requested-with", "XMLHttpRequest");

            result = http.GetHtml(item);

            if (result.Html.ToString().IndexOf("登录成功") > 0)
            {
                item = new HttpItem
                {
                    URL = "https://www.xm-l-tax.gov.cn/login/gdslhrz.do", //URL     必需项
                    Method = "post", //URL     可选项 默认为Get
                    Referer = "https://www.xm-l-tax.gov.cn/", //来源URL     可选项
                    Cookie = scookie,
                    IsToLower = false, //得到的HTML代码是否转成小写     可选项默认转小写
                    ContentType = "application/x-www-form-urlencoded; charset=UTF-8", //返回类型    可选项有默认值
                    UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                    Postdata =
                        "loginId=" + strName + "&userPassword=" + strPass + "&checkCode=" + strText, //Post数据
                };
                item.Header.Add("x-requested-with", "XMLHttpRequest");
                result = http.GetHtml(item);

                item = new HttpItem
                {
                    URL = "https://www.xm-l-tax.gov.cn/login/opener.do", //URL     必需项
                    Referer = "https://www.xm-l-tax.gov.cn/", //来源URL     可选项
                    Cookie = scookie,
                    UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                };
                result = http.GetHtml(item);
                item = new HttpItem
                {
                    URL = "https://www.xm-l-tax.gov.cn/nsfwHome/index.do", //URL     必需项
                    Cookie = scookie,
                    UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                };
                result = http.GetHtml(item);
                item = new HttpItem
                {
                    URL = "https://www.xm-l-tax.gov.cn/xxtx/index.do", //URL     必需项
                    Cookie = scookie,
                    UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                };
                result = http.GetHtml(item);
                item = new HttpItem
                {
                    URL = "https://www.xm-l-tax.gov.cn/nsfwHome/index.do?menuid=zhcx", //URL     必需项
                    Referer = "https://www.xm-l-tax.gov.cn/xxtx/index.do", //来源URL     可选项
                    Cookie = scookie,
                    UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                };
                result = http.GetHtml(item);
                item = new HttpItem
                {
                    URL = "https://www.xm-l-tax.gov.cn/sssq/swdjxxwh/swdjxxview.do", //URL     必需项
                    Referer = "https://www.xm-l-tax.gov.cn/nsfwHome/index.do?menuid=zhcx", //来源URL     可选项
                    Cookie = scookie,
                    UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                };
                result = http.GetHtml(item);
                string html = result.Html;
                Regex reg = new Regex(@"<table[^>]*>[\s\S]*</table>",
                    RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.Compiled);
                html = reg.Match(html).Value;
                html =
                    Regex.Replace(html, @"&nbsp;", "",
                        RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.Compiled);
                html =
                    Regex.Replace(html, @"^\s+|(\>)\s+(\<)|\s+$", "$1$2",
                        RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.Compiled).Replace("\r", "").Replace("\n", "");
                
                reg = new Regex(@"<!--[\s\S]*?-->", RegexOptions.IgnoreCase);
                html = reg.Replace(html, "");
                html = html.Replace("</td>", "</td>\t").Replace("</tr>", "</tr>\n");
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(html);
                HtmlNode node = doc.DocumentNode;
                html = node.InnerText;
                DataTable dt = new DataTable();
                string[] strRows = html.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries); //解析成行的字符串数组
                for (int rowIndex = 0; rowIndex < strRows.Length; rowIndex++) //行的字符串数组遍历
                {
                    string vsRow = strRows[rowIndex]; //取行的字符串
                    string[] vsColumns = vsRow.Split(new string[] { "\t" }, StringSplitOptions.RemoveEmptyEntries); //解析成字段数组
                    int k = vsColumns.Length - dt.Columns.Count + 1;
                    for (int ii = 1; ii <= k; ii++)
                    {
                        dt.Columns.Add();
                    }
                    dt.Rows.Add(vsColumns);
                }
                string[,] dts = new string[dt.Rows.Count, dt.Columns.Count];
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        dts[i, j] = dt.Rows[i][j].ToString();
                    }
                }
                Range rng = WorkingPaper.Wb.Worksheets["地税、基本情况"].Range["A1"].Resize[dts.GetLength(0), dts.GetLength(1)];
                rng.Value2 = dts;
                WorkingPaper.Wb.Worksheets["地税、基本情况"].Range["H6"].Value =
                    CU.Zifu(WorkingPaper.Wb.Worksheets["地税、基本情况"].Range["H6"].Value2).Trim();
                return true;
            }
            else
            {
                MessageBox.Show(result.Html);
                return false;
            }
        }

        private Boolean 国税信息(string strName, string strPass)
        {
            string strText, scookie;
            HttpHelper http = new HttpHelper();
            HttpItem item = new HttpItem()
            {
                URL = "http://wsbsdt.xm-n-tax.gov.cn:8001/", //URL     必需项
                IsToLower = false, //得到的HTML代码是否转成小写     可选项默认转小写
                UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
            };
            HttpResult result = http.GetHtml(item);
            scookie = result.Cookie;

            item = new HttpItem
            {
                URL = "http://wsbsdt.xm-n-tax.gov.cn:8001/bsfw/login/checkcode.do?r=Math.random()&ct=bsfw",
                //URL     必需项
                Referer = "http://wsbsdt.xm-n-tax.gov.cn:8001/", //来源URL     可选项
                IsToLower = false, //得到的HTML代码是否转成小写     可选项默认转小写
                Cookie = scookie,
                ResultType = ResultType.Byte, //返回数据类型，是Byte还是String
                UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
            };
            result = http.GetHtml(item);

            //把得到的Byte转成图片
            Image img = byteArrayToImage(result.ResultByte);
            验证码 pic = new 验证码(img, "国税验证码");
            pic.StartPosition = FormStartPosition.CenterParent;
            pic.ShowDialog();
            strText = pic.pictext;

            item = new HttpItem
            {
                URL = "http://wsbsdt.xm-n-tax.gov.cn:8001/bsfw/login/checkAndLogin.do", //URL     必需项
                Method = "post", //URL     可选项 默认为Get
                Referer = "http://wsbsdt.xm-n-tax.gov.cn:8001/", //来源URL     可选项
                Cookie = scookie,
                IsToLower = false, //得到的HTML代码是否转成小写     可选项默认转小写
                ContentType = "application/json;charset=UTF-8", //返回类型    可选项有默认值
                UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                Postdata =
                    @"{""login"":{""nsrsbh"":""" + strName + @""",""yhmm"":""" + strPass + @""",""checkCode"":""" +
                    strText + @""",""loginType"":""bsfw""}}", //Post数据
            };
            item.Header.Add("x-requested-with", "XMLHttpRequest");

            result = http.GetHtml(item);

            if (result.Html.ToString().IndexOf("登录成功") > 0)
            {
                item = new HttpItem
                {
                    URL = "http://wsbsdt.xm-n-tax.gov.cn:8001/bsfw/home/index.do", //URL     必需项
                    Referer = "http://wsbsdt.xm-n-tax.gov.cn:8001/", //来源URL     可选项
                    Cookie = scookie,
                    UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                };
                result = http.GetHtml(item);
                item = new HttpItem
                {
                    URL = "http://wsbsdt.xm-n-tax.gov.cn:8001/bsfw/home/sscx_index.do", //URL     必需项
                    Referer = "http://wsbsdt.xm-n-tax.gov.cn:8001/bsfw/home/index.do", //来源URL     可选项
                    Cookie = scookie,
                    UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                };
                result = http.GetHtml(item);
                item = new HttpItem
                {
                    URL = "http://wsbsdt.xm-n-tax.gov.cn:8001/bsfw/nsrgl/queryNsrjbxx.do", //URL     必需项
                    Referer = "http://wsbsdt.xm-n-tax.gov.cn:8001/bsfw/home/sscx_index.do", //来源URL     可选项
                    Cookie = scookie,
                    UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                };
                result = http.GetHtml(item);
                string html = result.Html;
                Regex reg = new Regex(@"<table[^>]*>[\s\S]*</table>",
                    RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.Compiled);
                html = reg.Match(html).Value;
                html =
                    Regex.Replace(html, @"&nbsp;", "",
                        RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.Compiled);
                html =
                    Regex.Replace(html, @"^\s+|(\>)\s+(\<)|\s+$", "$1$2",
                        RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.Compiled).Replace("\r", "").Replace("\n", "");
                HtmlTableService ht = new HtmlTableService();
                string[,] dt = ht.ToArray(html, Encoding.UTF8);
                Range rng = WorkingPaper.Wb.Worksheets["地税、基本情况"].Range["W1"].Resize[dt.GetLength(0), dt.GetLength(1)];
                rng.Value2 = dt;

                return true;
            }
            else
            {
                MessageBox.Show(result.Html);
                return false;
            }
        }

        private string base64(string pass)  //base64加密
        {
            HttpHelper http = new HttpHelper();
            HttpItem item = new HttpItem()
            {
                URL = "https://www.xm-l-tax.gov.cn/res/js/Base64.js", //URL     必需项
                IsToLower = false, //得到的HTML代码是否转成小写     可选项默认转小写
                UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
            };
            HttpResult result = http.GetHtml(item);
            return ExecuteScript("base64encode(\"" + pass + "\")", result.Html);
        }

        private string md5(string pass)
        {
            HttpHelper http = new HttpHelper();
            HttpItem item = new HttpItem()
            {
                URL = "https://www.xm-l-tax.gov.cn/res/js/md5.js", //URL     必需项
                IsToLower = false, //得到的HTML代码是否转成小写     可选项默认转小写
                UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
            };
            HttpResult result = http.GetHtml(item);
            return ExecuteScript("hex_md5(\"" + pass + "\")", result.Html);
        }

        private string ExecuteScript(string sExpression, string sCode)
        {
            MSScriptControl.ScriptControl scriptControl = new MSScriptControl.ScriptControl();
            scriptControl.UseSafeSubset = true;
            scriptControl.Language = "JScript";
            scriptControl.AddCode(sCode);
            try
            {
                string str = scriptControl.Eval(sExpression).ToString();
                return str;
            }
            catch (Exception ex)
            {
                string str = ex.Message;
            }
            return null;
        }

        private Image byteArrayToImage(byte[] Bytes)
        {
            MemoryStream ms = new MemoryStream(Bytes);
            Image outputImg = Image.FromStream(ms);
            return outputImg;
        }

        /// <summary>
        /// 通过知乎专栏获取版本号
        /// </summary>
        /// <param name="sUrl">专栏地址</param>
        /// <returns>版本号</returns>
        public static string 获取版本号(string sUrl)
        {
            HttpHelper http = new HttpHelper();
            HttpItem item = new HttpItem()
            {
                URL = sUrl, //URL     必需项
                IsToLower = false, //得到的HTML代码是否转成小写     可选项默认转小写
                UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
            };
            HttpResult result = http.GetHtml(item);
            string str = result.Html;
            str = Regex.Match(str, @"(?<=@@)([\S\s]*?)(?=@@)").ToString();
            if (str == "")
            {
                str = "获取失败";
            }
            return str;
        }
        #endregion
    }
}
