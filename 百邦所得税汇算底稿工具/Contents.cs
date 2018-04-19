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
            treeView2.Nodes.Clear();
            treeView2.Nodes.Add("刷新底稿目录");
        }
        
        void 树状图()
        {

            if (WorkingPaper.版本号 == 2016 || WorkingPaper.版本号 == 2017)
            {
                treeView1.Nodes.Clear();
                string[] Text, Tag;
                TreeNode Tn;

                Tn = treeView1.Nodes.Add("综合类底稿");
                Text = Properties.Resources.Text1.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                Tag = Properties.Resources.Tag1.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < Text.Length; i++)
                {
                    TreeNode tn = new TreeNode
                    {
                        Tag = Tag[i],
                        Text = Text[i]
                    };
                    Tn.Nodes.Add(tn);
                }
                Tn = treeView1.Nodes.Add("调整类底稿");
                Text = Properties.Resources.Text2.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                Tag = Properties.Resources.Tag2.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < Text.Length; i++)
                {
                    TreeNode tn = new TreeNode
                    {
                        Tag = Tag[i],
                        Text = Text[i]
                    };
                    Tn.Nodes.Add(tn);
                }
                Tn = treeView1.Nodes.Add("资产类底稿");
                Text = Properties.Resources.Text5.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                Tag = Properties.Resources.Tag5.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < Text.Length; i++)
                {
                    TreeNode tn = new TreeNode
                    {
                        Tag = Tag[i],
                        Text = Text[i]
                    };
                    Tn.Nodes.Add(tn);
                }
                Tn = treeView1.Nodes.Add("负债及权益类底稿");
                Text = Properties.Resources.Text6.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                Tag = Properties.Resources.Tag6.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < Text.Length; i++)
                {
                    TreeNode tn = new TreeNode
                    {
                        Tag = Tag[i],
                        Text = Text[i]
                    };
                    Tn.Nodes.Add(tn);
                }
                Tn = treeView1.Nodes.Add("损益类底稿");
                Text = Properties.Resources.Text3.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                Tag = Properties.Resources.Tag3.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < Text.Length; i++)
                {
                    TreeNode tn = new TreeNode
                    {
                        Tag = Tag[i],
                        Text = Text[i]
                    };
                    Tn.Nodes.Add(tn);
                }
                Tn = treeView1.Nodes.Add("审核报告");
                Text = Properties.Resources.Text4.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                Tag = Properties.Resources.Tag4.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < Text.Length; i++)
                {
                    TreeNode tn = new TreeNode
                    {
                        Tag = Tag[i],
                        Text = Text[i]
                    };
                    Tn.Nodes.Add(tn);
                }
            }
            else
            {

                treeView1.Nodes.Clear();
                string[] Text, Tag;
                TreeNode Tn;

                Tn = treeView1.Nodes.Add("Integrated综合类底稿");
                Text = Properties.Resources.Tag81.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                Tag = Properties.Resources.Tag81.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < Text.Length; i++)
                {
                    TreeNode tn = new TreeNode
                    {
                        Tag = Tag[i],
                        Text = Text[i]
                    };
                    Tn.Nodes.Add(tn);
                }
                Tn = treeView1.Nodes.Add("Basic基本资料类底稿");
                Text = Properties.Resources.Tag82.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                Tag = Properties.Resources.Tag82.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < Text.Length; i++)
                {
                    TreeNode tn = new TreeNode
                    {
                        Tag = Tag[i],
                        Text = Text[i]
                    };
                    Tn.Nodes.Add(tn);
                }
                Tn = treeView1.Nodes.Add("Assets资产类底稿");
                Text = Properties.Resources.Tag83.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                Tag = Properties.Resources.Tag83.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < Text.Length; i++)
                {
                    TreeNode tn = new TreeNode
                    {
                        Tag = Tag[i],
                        Text = Text[i]
                    };
                    Tn.Nodes.Add(tn);
                }
                Tn = treeView1.Nodes.Add("Liabilities and Rights负债及权益类底稿");
                Text = Properties.Resources.Tag84.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                Tag = Properties.Resources.Tag84.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < Text.Length; i++)
                {
                    TreeNode tn = new TreeNode
                    {
                        Tag = Tag[i],
                        Text = Text[i]
                    };
                    Tn.Nodes.Add(tn);
                }
                Tn = treeView1.Nodes.Add("Gains and Losses损益类底稿");
                Text = Properties.Resources.Tag85.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                Tag = Properties.Resources.Tag85.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < Text.Length; i++)
                {
                    TreeNode tn = new TreeNode
                    {
                        Tag = Tag[i],
                        Text = Text[i]
                    };
                    Tn.Nodes.Add(tn);
                }
                Tn = treeView1.Nodes.Add("Adjustment调整类底稿");
                Text = Properties.Resources.Tag86.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                Tag = Properties.Resources.Tag86.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < Text.Length; i++)
                {
                    TreeNode tn = new TreeNode
                    {
                        Tag = Tag[i],
                        Text = Text[i]
                    };
                    Tn.Nodes.Add(tn);
                }
                Tn = treeView1.Nodes.Add("Declaration报告及申报表");
                Text = Properties.Resources.Tag87.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                Tag = Properties.Resources.Tag87.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < Text.Length; i++)
                {
                    TreeNode tn = new TreeNode
                    {
                        Tag = Tag[i],
                        Text = Text[i]
                    };
                    Tn.Nodes.Add(tn);
                }
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
                    if (splitContainer1.Height - 100 >= splitContainer1.Panel1MinSize)
                    {
                        splitContainer1.SplitterDistance = 200;
                    }

                    break;
                case "税金申报明细":
                    //tP税费.Parent= tabControl1;
                    splitContainer1.Panel1Collapsed = false;
                    splitContainer3.Panel2Collapsed = false;
                    splitContainer3.Panel1Collapsed = true;
                    splitContainer4.Panel1Collapsed = false;
                    splitContainer4.Panel2Collapsed = true;
                    if (splitContainer1.Height - 100 >= splitContainer1.Panel1MinSize)
                    {
                        splitContainer1.SplitterDistance = 150;
                    }

                    groupBox1.Text = "税费测算";
                    break;
                case "基本情况":
                    //tP税费.Parent= tabControl1;
                    splitContainer1.Panel1Collapsed = false;
                    splitContainer3.Panel2Collapsed = false;
                    splitContainer3.Panel1Collapsed = true;
                    splitContainer4.Panel2Collapsed = false;
                    splitContainer4.Panel1Collapsed = true;
                    if (splitContainer1.Height-100>=splitContainer1.Panel1MinSize)
                    {
                        splitContainer1.SplitterDistance = 100;

                    }
                    groupBox1.Text = "基本情况";
                    break;
                default:
                    splitContainer1.Panel1Collapsed = true;
                    break;
            }
        }

        private void Button2_Click(object sender, EventArgs e)//期间费用
        {
            CU.文件判断();
            if (WorkingPaper.Wb.ActiveSheet.Name == "余额表")
            {
                QJFY qj = new QJFY(WorkingPaper.Wb.ActiveSheet);
                qj.ShowDialog();
            }

        }

        private void Button1_Click(object sender, EventArgs e)//2017报表填写
        {

            if (WorkingPaper.版本号 == 2018)
            {

                Codes2018.报表填写();
            }
            else
            {

                _2017Olds.报表填写();
            }

        }

        private void Button3_Click(object sender, EventArgs e)//2017底稿填写
        {

            if (WorkingPaper.版本号 == 2018)
            {

                Codes2018.底稿填写();
            }
            else
            {
                _2017Olds.底稿填写();
            }
        }      

        private void TreeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)//双击树状图
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

        private void Button4_Click(object sender, EventArgs e)//2017税费填写
        {
            if (WorkingPaper.版本号 == 2018)
            {

                Codes2018.税费填写();
            }
            else
            {
                _2017Olds.税费填写();
            }
        }

        private void TreeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        private void GroupBox1_Enter(object sender, EventArgs e)
        {

        }

        #region 基本情况
        private void btn基础信息_Click(object sender, EventArgs e)
        {
            if (!CU.文件判断())
                return;

            if (WorkingPaper.版本号 == 2018)
            {

                string name1, pass1;
                name1 = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["F3"].Value2);
                pass1 = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["H3"].Value2);
                if (name1 == "" || pass1 == "")
                    MessageBox.Show("国税用户名密码未填写，请填写[基本情况].[F3,H3]后重试！");
                else if (国税信息(name1, base64(pass1)))
                    MessageBox.Show("国税信息抓取成功！");
                else
                    MessageBox.Show("国税抓取失败！");
                name1 = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["F4"].Value2);
                pass1 = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["H4"].Value2);
                if (name1 == "" || pass1 == "")
                    MessageBox.Show("地税用户名密码未填写，请填写[基本情况].[F4,H4]后重试！");
                else if (地税信息(name1, pass1))
                    MessageBox.Show("地税信息抓取成功！");
                else
                    MessageBox.Show("地税抓取失败！");
            }
            else
            {

                string name1, pass1;
                name1 = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["B49"].Value2);
                pass1 = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["D49"].Value2);
                if (name1 == "" || pass1 == "")
                    MessageBox.Show("国税用户名密码未填写，请填写[基本情况].[B49,D49]后重试！");
                else if (国税信息(name1, base64(pass1)))
                    MessageBox.Show("国税信息抓取成功！");
                else
                    MessageBox.Show("国税抓取失败！");
                name1 = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["B50"].Value2);
                pass1 = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["D50"].Value2);
                if (name1 == "" || pass1 == "")
                    MessageBox.Show("地税用户名密码未填写，请填写[基本情况].[B50,D50]后重试！");
                else if (地税信息(name1, pass1))
                    MessageBox.Show("地税信息抓取成功！");
                else
                    MessageBox.Show("地税抓取失败！");
            }

        }
        #endregion

        #region 申报数据获取

        private void btnDP_Click(object sender, EventArgs e)
        {
            string strText, scookie, strName, strPass,year,ZSXMGS,ZSPMGS;
            if (!CU.文件判断())
                return;

            if (WorkingPaper.版本号 == 2018)
            {

                strName = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["F4"].Value2);
                strPass = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["H4"].Value2);
                year = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["F6"].Value2);
                if (strName == "" || strPass == "")
                {
                    MessageBox.Show("地税用户名和密码未填写，请填写[基本情况].[F4,H4]后重试！");
                    return;
                }

                ZSXMGS = "=VLOOKUP(RC[-11],辅助表!C[4]:C[5],2,0)";
                ZSPMGS = "=VLOOKUP(RC[-11],辅助表!C[5]:C[6],2,0)";
            }
            else
            {

                strName = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["B50"].Value2);
                strPass = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["D50"].Value2);

                year = CU.Zifu(WorkingPaper.Wb.Worksheets["基本情况"].Range["B4"].Value2);
                if (strName == "" || strPass == "")
                {
                    MessageBox.Show("地税用户名和密码未填写，请填写[基本情况].[B50,D50]后重试！");
                    return;
                }
                ZSXMGS = "=VLOOKUP(RC[-11],首页!C[-6]:C[-5],2,0)";
                ZSPMGS = "=VLOOKUP(RC[-11],首页!C[-5]:C[-4],2,0)";
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
            验证码 pic = new 验证码(img, "地税验证码")
            {
                StartPosition = FormStartPosition.CenterParent
            };
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
                rng.FormulaR1C1 = ZSXMGS;
                rng = WorkingPaper.Wb.Worksheets["税金申报明细"].Range["N2"].Resize[dr.GetLength(0), 1];
                rng.FormulaR1C1 = ZSPMGS;
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
            验证码 pic = new 验证码(img, "地税验证码")
            {
                StartPosition = FormStartPosition.CenterParent
            };
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
            验证码 pic = new 验证码(img, "国税验证码")
            {
                StartPosition = FormStartPosition.CenterParent
            };
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

        private string Md5(string pass)
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
            MSScriptControl.ScriptControl scriptControl = new MSScriptControl.ScriptControl
            {
                UseSafeSubset = true,
                Language = "JScript"
            };
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

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void treeView2_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        private void treeView2_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            树状图();
        }
        
    }
}
