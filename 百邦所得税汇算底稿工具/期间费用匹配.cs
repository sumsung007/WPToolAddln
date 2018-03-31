using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace 百邦所得税汇算底稿工具
{
    public partial class QJFY : Form
    {

        Worksheet Sh;
        string i, j,k,l;
        int N;
        object[,] Yeb,Kmlb;

        public QJFY(Worksheet sh)
        {
            InitializeComponent();
            Sh = sh;
            ListRefresh();
            listView2.ItemDrag += new ItemDragEventHandler(ListView2_ItemDrag);
            listView1.DragEnter += new DragEventHandler(ListView1_DragEnter);
            listView1.DragDrop += new DragEventHandler(ListView1_DragDrop);
            //listView1.DragOver += new DragEventHandler(ListView1_DragOver);
            this.FormClosing += QJFY_FormClosing;
        }

        private void QJFY_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (WorkingPaper.版本号 == 2018)
            {
                Sh.Range["J5:J" + N].Value2 = Kmlb;
            }
            else
            {
                Sh.Range["J2:J" + N].Value2 = Kmlb;
            }
        }

        /*private void ListView1_DragOver(object sender, DragEventArgs e)
        {

            System.Drawing.Point ptScreen = new System.Drawing.Point(e.X, e.Y);
            System.Drawing.Point pt = listView1.PointToClient(ptScreen);
            ListViewItem item = listView1.GetItemAt(pt.X, pt.Y);
            if (item != null)
            {
                item.Selected=true;
            }
            }
        */
        private void ListView2_ItemDrag(object sender, ItemDragEventArgs e)
        {
            i = listView2.SelectedItems[0].SubItems[1].Text;
            j = listView2.SelectedItems[0].SubItems[2].Text;
            k = listView2.SelectedItems[0].SubItems[3].Text;
            l = e.Item.ToString();
            l = l.Substring(l.IndexOf("、")+1);
            l = l.Substring(0, l.Length - 1);
            this.DoDragDrop(e.Item,DragDropEffects.All);
        }
        private void ListView1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = e.AllowedEffect;

        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("该操作将清除所有匹配，不可恢复，请谨慎处理！是否继续清除？", "警告",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                Sh.Columns[10].Clear();
                ListRefresh();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string T,m,G;
            if (MessageBox.Show("该操作将清除所有匹配，不可恢复，请谨慎处理！是否继续清除？", "警告",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                Sh.Columns[10].Clear();
                ListRefresh();
                foreach (ListViewItem item in listView1.Items)
                {
                    if (item.SubItems[5].Text == "0")
                    {

                        m = "其他";
                        T = item.SubItems[1].Text;
                        G = item.Group.Header;
                        if (G == "财务费用")
                        {

                            if (T.Contains("手续"))
                                m = "金融机构手续费";
                            if ((T.Contains("利息费")) || (T.Contains("利息支出")))
                                m = "利息支出";
                            if (T.Contains("汇兑损失"))
                                m = "汇兑损失";
                            if (T.Contains("现金折扣"))
                                m = "现金折扣";
                            if (T.Contains("利息收入"))
                                m = "利息收入";
                            if (T.Contains("汇兑收益"))
                                m = "汇兑收益";
                            if (T.Contains("佣金"))
                                m = "佣金和手续费";
                        }
                        else
                        {
                            if ((T.Contains("工资")) || (T.Contains("薪金")) || (T.Contains("薪酬")))
                                m = "工资薪金支出";
                            if (T.Contains("福利费"))
                                m = "职工福利费支出";
                            if (T.Contains("教育"))
                                m = "职工教育经费支出";
                            if (T.Contains("工会"))
                                m = "工会经费支出";
                            if ((T.Contains("招待")) || (T.Contains("应酬")))
                                m = "业务招待费支出";
                            if ((T.Contains("广告")) || (T.Contains("宣传")))
                                m = "广告费和业务宣传费支出";
                            if ((T.Contains("捐赠")) || (T.Contains("捐款")))
                                m = "捐赠支出";
                            if (T.Contains("公积"))
                                m = "住房公积金";
                            if (Regex.IsMatch(T, "罚款|罚金|没收"))
                                m = "罚金、罚款和被没收财物的损失";
                            if (T.Contains("滞纳"))
                                m = "税收滞纳金";
                            if (T.Contains("赞助"))
                                m = "赞助支出";
                            if (Regex.IsMatch(T, "社.*保"))
                                m = "各类基本社会保障性缴款";
                            if (Regex.IsMatch(T, "补充养老|补充医疗|年金"))
                                m = "补充养老险、补充医疗保险";
                            if (Regex.IsMatch(T, "(财.*损)|(资.*损)"))
                                m = "财产损失";
                            if (T.Contains("折旧"))
                                m = "折旧";
                            if (Regex.IsMatch(T, "房租|租金|租赁费"))
                                m = "租金";
                            if (Regex.IsMatch(T, "辞.*福"))
                                m = "辞退福利";
                            if (Regex.IsMatch(T, "无形.*摊销"))
                                m = "无形资产摊销";
                            if (Regex.IsMatch(T, "长期.*摊销"))
                                m = "长期待摊费用摊销";
                            if (T.Contains("印花税"))
                                m = "印花税";
                            if (T.Contains("房产税"))
                                m = "房产税";
                            if (T.Contains("土地使用税"))
                                m = "土地使用税";
                            if (T.Contains("车船使用税"))
                                m = "车船使用税";
                            if (T.Contains("劳务"))
                                m = "劳务费";
                            if ((T.Contains("咨询")) || (T.Contains("顾问")))
                                m = "咨询顾问费";
                            if (T.Contains("佣金"))
                                m = "佣金和手续费";
                            if (T.Contains("办公"))
                                m = "办公费";
                            if (T.Contains("董事会"))
                                m = "董事会费";
                            if (Regex.IsMatch(T, "诉讼.*费"))
                                m = "诉讼费";
                            if (T.Contains("差旅"))
                                m = "差旅费";
                            if (Regex.IsMatch(T, "(财.*保险)|(身.*保险)|(责.*保险)|(综.*保险)|(危.*保险)|(交.*保险)"))
                                m = "保险费";
                            if (Regex.IsMatch(T, "(运输.*费)|(仓储.*费)"))
                                m = "运输、仓储费";
                            if (T.Contains("修理"))
                                m = "修理费";
                            if (T.Contains("包装"))
                                m = "包装费";
                            if (Regex.IsMatch(T, "技术.*转让"))
                                m = "技术转让费";
                            if ((T.Contains("研究")) || (T.Contains("研发")))
                                m = "研究费用";
                            if (WorkingPaper.版本号 == 2018 && T.Contains("党组织工作经费") && G == "管理费用")
                                m = "党组织工作经费";
                        }
                        item.SubItems[3].Text = m;
                        Kmlb[Convert.ToInt16(item.SubItems[4].Text), 1] = G + "-" + m;
                    }
                }
            }
        }

        private void ListView1_DragDrop(object sender, DragEventArgs e)
        {
            System.Drawing.Point ptScreen = new System.Drawing.Point(e.X, e.Y);
            System.Drawing.Point pt = listView1.PointToClient(ptScreen);
            ListViewItem item = listView1.GetItemAt(pt.X, pt.Y);
            if (item != null)
            {
                if(item.SubItems[5].Text == "0")
                    switch (item.Group.Header.ToString())
                    {
                        case "销售费用":
                            if (i == "√")
                            {
                                item.SubItems[3].Text = l;
                                Kmlb[Convert.ToInt16(item.SubItems[4].Text), 1] = "销售费用-" + l;
                            }
                            break;
                        case "管理费用":
                            if (j == "√")
                            {
                                if ((! (WorkingPaper.版本号 == 2018)) && l== "党组织工作经费")
                                    break;
                                item.SubItems[3].Text = l;
                                Kmlb[Convert.ToInt16(item.SubItems[4].Text), 1] = "管理费用-" + l;
                            }
                            break;
                        default:
                            if (k == "√")
                            {
                                item.SubItems[3].Text = l;
                                Kmlb[Convert.ToInt16(item.SubItems[4].Text), 1] = "财务费用-" + l;
                            }
                            break;
                    }
            }
        }

        //
        void ListRefresh()
        {
            N = Sh.Cells[Sh.UsedRange.Rows.Count + 1, 2].End[XlDirection.xlUp].Row;
            if (N>1)
            {
                int yyfy, glfy, cwfy, i,k;
                string yykm, glkm, cwkm;
                this.listView1.BeginUpdate();
                this.listView1.Items.Clear();
                if (WorkingPaper.版本号 == 2018)
                {
                    Yeb = Sh.Range["A5:H" + N].Value2;    //余额表
                    Kmlb = Sh.Range["J5:J" + N].Value2;   //科目类别
                    yyfy = Convert.ToInt16(CU.Shuzi(Sh.Range["K1"].Value));      //营业费用
                    glfy = Convert.ToInt16(CU.Shuzi(Sh.Range["L1"].Value));      //管理费用
                    cwfy = Convert.ToInt16(CU.Shuzi(Sh.Range["M1"].Value));      //财务费用
                    k = 3;//偏移
                }
                else
                {
                    Yeb = Sh.Range["A2:H" + N].Value2;    //余额表
                    Kmlb = Sh.Range["J2:J" + N].Value2;   //科目类别
                    yyfy = Convert.ToInt16(CU.Shuzi(Sh.Range["N11"].Value));      //营业费用
                    glfy = Convert.ToInt16(CU.Shuzi(Sh.Range["N10"].Value));      //管理费用
                    cwfy = Convert.ToInt16(CU.Shuzi(Sh.Range["N13"].Value));      //财务费用   
                    k = 0;//偏移
                }
                if (yyfy != 0)//营业费用
                {
                    yykm = CU.Zifu(Yeb[yyfy - 1, 1]);
                    if ((yyfy == N-k) || (!CU.Zifu(Yeb[yyfy, 1]).Contains(yykm)))
                    {
                        ListViewItem lvi = new ListViewItem();
                        lvi.Group = listView1.Groups["yyfy"];
                        lvi.Text = CU.Zifu(Yeb[yyfy - 1, 1]);
                        lvi.SubItems.Add(CU.Zifu(Yeb[yyfy - 1, 2]));
                        lvi.SubItems.Add(CU.Shuzi(Yeb[yyfy - 1, 5]).ToString("N"));
                        lvi.SubItems.Add(Kmlb[yyfy - 1, 1] == null ? "" : CU.Zifu(Kmlb[yyfy - 1, 1]).
                            Substring(CU.Zifu(Kmlb[yyfy - 1, 1]).IndexOf("-") + 1));
                        lvi.SubItems.Add((yyfy - 1).ToString());
                        lvi.SubItems.Add("0");
                        this.listView1.Items.Add(lvi);
                    }
                    else
                    {
                        i = yyfy;
                        do
                        {
                            ListViewItem lvi = new ListViewItem();
                            lvi.Group = listView1.Groups["yyfy"];
                            lvi.Text = CU.Zifu(Yeb[i, 1]);
                            lvi.SubItems.Add(CU.Zifu(Yeb[i, 2]));
                            lvi.SubItems.Add(CU.Shuzi(Yeb[i, 5]).ToString("N"));
                            lvi.SubItems.Add(Kmlb[i, 1] == null ? "" : CU.Zifu(Kmlb[i, 1]).
                                Substring(CU.Zifu(Kmlb[i, 1]).IndexOf("-") + 1));
                            lvi.SubItems.Add(i.ToString());
                            if ((i == N - k-1) || (!CU.Zifu(Yeb[i + 1, 1]).Contains(CU.Zifu(Yeb[i, 1]))))
                                lvi.SubItems.Add("0");
                            else
                                lvi.SubItems.Add("1");
                            this.listView1.Items.Add(lvi);
                            i++;
                        } while ((i < N-k) && (CU.Zifu(Yeb[i, 1]).Contains(yykm)));
                    }
                }

                if (glfy != 0)//管理费用
                {
                    glkm = CU.Zifu(Yeb[glfy - 1, 1]);
                    if ((glfy == N-k) || (!CU.Zifu(Yeb[glfy, 1]).Contains(glkm)))
                    {
                        ListViewItem lvi = new ListViewItem();
                        lvi.Group = listView1.Groups["glfy"];
                        lvi.Text = CU.Zifu(Yeb[glfy - 1, 1]);
                        lvi.SubItems.Add(CU.Zifu(Yeb[glfy - 1, 2]));
                        lvi.SubItems.Add(CU.Shuzi(Yeb[glfy - 1, 5]).ToString("N"));
                        lvi.SubItems.Add(Kmlb[glfy - 1, 1] == null ? "" : CU.Zifu(Kmlb[glfy - 1, 1]).
                            Substring(CU.Zifu(Kmlb[glfy - 1, 1]).IndexOf("-") + 1));
                        lvi.SubItems.Add((glfy - 1).ToString());
                        lvi.SubItems.Add("0");
                        this.listView1.Items.Add(lvi);
                    }
                    else
                    {
                        i = glfy;
                        do
                        {
                            ListViewItem lvi = new ListViewItem();
                            lvi.Group = listView1.Groups["glfy"];
                            lvi.Text = CU.Zifu(Yeb[i, 1]);
                            lvi.SubItems.Add(CU.Zifu(Yeb[i, 2]));
                            lvi.SubItems.Add(CU.Shuzi(Yeb[i, 5]).ToString("N"));
                            lvi.SubItems.Add(Kmlb[i, 1] == null ? "" : CU.Zifu(Kmlb[i, 1]).
                                Substring(CU.Zifu(Kmlb[i, 1]).IndexOf("-") + 1));
                            lvi.SubItems.Add(i.ToString());
                            if ((i == N - k-1) || (!CU.Zifu(Yeb[i + 1, 1]).Contains(CU.Zifu(Yeb[i, 1]))))
                                lvi.SubItems.Add("0");
                            else
                                lvi.SubItems.Add("1");
                            this.listView1.Items.Add(lvi);
                            i++;
                        } while ((i < N-k) && (CU.Zifu(Yeb[i, 1]).Contains(glkm)));
                    }
                }

                if (cwfy != 0)//财务费用
                {
                    cwkm = CU.Zifu(Yeb[cwfy - 1, 1]);
                    if ((cwfy == N-k) || (!CU.Zifu(Yeb[cwfy, 1]).Contains(cwkm)))
                    {
                        ListViewItem lvi = new ListViewItem();
                        lvi.Group = listView1.Groups["cwfy"];
                        lvi.Text = CU.Zifu(Yeb[cwfy - 1, 1]);
                        lvi.SubItems.Add(CU.Zifu(Yeb[cwfy - 1, 2]));
                        lvi.SubItems.Add(CU.Shuzi(Yeb[cwfy - 1, 5]).ToString("N"));
                        lvi.SubItems.Add(Kmlb[cwfy - 1, 1] == null ? "" : CU.Zifu(Kmlb[cwfy - 1, 1]).
                            Substring(CU.Zifu(Kmlb[cwfy - 1, 1]).IndexOf("-") + 1));
                        lvi.SubItems.Add((cwfy - 1).ToString());
                        lvi.SubItems.Add("0");
                        this.listView1.Items.Add(lvi);
                    }
                    else
                    {
                        i = cwfy;
                        do
                        {
                            ListViewItem lvi = new ListViewItem();
                            lvi.Group = listView1.Groups["cwfy"];
                            lvi.Text = CU.Zifu(Yeb[i, 1]);
                            lvi.SubItems.Add(CU.Zifu(Yeb[i, 2]));
                            lvi.SubItems.Add(CU.Shuzi(Yeb[i, 5]).ToString("N"));
                            lvi.SubItems.Add(Kmlb[i, 1] == null ? "" : CU.Zifu(Kmlb[i, 1]).
                                Substring(CU.Zifu(Kmlb[i, 1]).IndexOf("-") + 1));
                            lvi.SubItems.Add(i.ToString());
                            if ((i == N - k-1) || (!CU.Zifu(Yeb[i + 1, 1]).Contains(CU.Zifu(Yeb[i, 1]))))
                                lvi.SubItems.Add("0");
                            else
                                lvi.SubItems.Add("1");
                            this.listView1.Items.Add(lvi);
                            i++;
                        } while ((i < N-k) && (CU.Zifu(Yeb[i, 1]).Contains(cwkm)));
                    }
                }
            }
            this.listView1.EndUpdate();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
