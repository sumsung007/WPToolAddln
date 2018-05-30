using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Management;
using System.Windows.Forms;

namespace 百邦所得税汇算底稿工具
{
    class CU
    {
        public static Boolean 事项说明()
        {

            if (WorkingPaper.OOO)
            {

                if (WorkingPaper.版本号 == 2018)
                {
                    WorkingPaper.Wb.Sheets["企业基本情况"].Range["$H$21:$H$128"].AutoFilter(Field: 1, Criteria1: "=1");
                }
                else
                {


                    Worksheet SH = WorkingPaper.Wb.Sheets["(二)附表-纳税调整额的审核"];
                    SH.Range["A7:E" + SH.Cells[SH.UsedRange.Rows.Count + 1, 1].End[XlDirection.xlUp].Row.ToString()]
                        .Value = "";
                    object[,] Nstz = WorkingPaper.Wb.Sheets["事项说明"].Range["B31:F81"].Value2;
                    List<string> Xiangmu = new List<string>();
                    List<double> Zhangzai = new List<double>();
                    List<double> Shuishou = new List<double>();
                    List<double> Tiaozheng = new List<double>();
                    //List<String> Yuanyin = new List<string>();

                    for (int tz = 1; tz <= 51; tz++)
                    {
                        if (CU.Shuzi(Nstz[tz, 5]) > 0)
                        {
                            Xiangmu.Add(CU.Zifu(Nstz[tz, 1]));
                            Zhangzai.Add(CU.Shuzi(Nstz[tz, 3]));
                            Shuishou.Add(CU.Shuzi(Nstz[tz, 4]));
                            Tiaozheng.Add(CU.Shuzi(Nstz[tz, 5]));
                            //Yuanyin.Add("税法规定");
                        }
                    }

                    Nstz = WorkingPaper.Wb.Sheets["事项说明"].Range["B85:F117"].Value2;
                    for (int tz = 1; tz <= 33; tz++)
                    {
                        if (CU.Shuzi(Nstz[tz, 5]) > 0)
                        {
                            Xiangmu.Add(CU.Zifu(Nstz[tz, 1]));
                            Zhangzai.Add(CU.Shuzi(Nstz[tz, 3]));
                            Shuishou.Add(CU.Shuzi(Nstz[tz, 4]));
                            Tiaozheng.Add(-CU.Shuzi(Nstz[tz, 5]));
                            //Yuanyin.Add("税法规定");
                        }
                    }

                    if (Xiangmu.Count > 0)
                    {
                        string[,] xiangmu = new string[Xiangmu.Count, 1];
                        double[,] jine = new double[Xiangmu.Count, 3];
                        string[,] yuanyin = new string[Xiangmu.Count, 1];
                        for (int k = 0; k < Xiangmu.Count; k++)
                        {
                            xiangmu[k, 0] = Xiangmu[k];
                            jine[k, 0] = Zhangzai[k];
                            jine[k, 1] = Shuishou[k];
                            jine[k, 2] = Tiaozheng[k];
                            yuanyin[k, 0] = "税法规定";
                        }

                        SH.Range["A7:A" + (Xiangmu.Count + 6).ToString()].Value2 = xiangmu;
                        SH.Range["B7:D" + (Xiangmu.Count + 6).ToString()].Value2 = jine;
                        SH.Range["E7:E" + (Xiangmu.Count + 6).ToString()].Value2 = yuanyin;
                    }
                }
            }
            return true;
        }

        //加密模块
        public static Boolean 授权检测()//读取注册表
        {
            if( Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\BaiBang", "Key", String.Empty) !=null)
                {
                if(Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\BaiBang", "Key", String.Empty).ToString() == 加密(机器码()))
                {
                    return true;
                }
                /*
                else
                {
                    if(DateTime.Now<Convert.ToDateTime("2016-07-31"))
                    {
                        
                        return true;
                    }
                }
                */
            }
            return false;
        }
        public static string 加密(String s)//转换为授权码
        {
            string str = "";
            for (int i = 0; i <= s.Length - 1; i++)
            {
                str = str + ((char)(((int)s[i] + i) % 61 + 65)).ToString();
            }
            return str;
        }
        static string GetCPUID()//获取CPUID
        {
            try
            {
                SelectQuery query = new SelectQuery("Win32_Processor");
                ManagementObjectSearcher search = new ManagementObjectSearcher(query);
                foreach (ManagementObject info in search.Get())
                {
                    if (info["ProcessorId"] != null)
                    {
                        return info["ProcessorId"].ToString();
                    }
                    else
                    {
                        return "";
                    }
                }
                return "";
            }
            catch (Exception)
            {
                return "";
            }
        }
        static string GetMainBoardID()//获取主板ID
        {
            try
            {
                ManagementObjectSearcher search = new ManagementObjectSearcher("Select * FROM Win32_BaseBoard");
                foreach (ManagementObject info in search.Get())
                {
                    if (info["Product"] != null)
                    {
                        return info["Product"].ToString();
                    }
                    else
                    {
                        return "";
                    }
                }
                return "";
            }
            catch (Exception)
            {
                return "";
            }
        }
        public static string 机器码()//获取机器码
        {
            string S1, S2;
            string str = "";
            S1 = GetCPUID().PadLeft(20, '0');
            S2 = GetMainBoardID().PadLeft(20, '0');
            for (int i = 0; i <= 19; i++)
            {
                str = str + ((char)(((int)S1[i] + (int)S2[19 - i]) % 61 + 65)).ToString().ToUpper();
            }
            return str;
        }
        //加密模块结束

        public static Boolean 文件判断()//判断是否为底稿文件
        {
            try
            {
                if (Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["基本情况"] != null)
                {
                    if (Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["基本情况"].range("B8").value
                        == "中汇百邦（厦门）税务师事务所有限公司" ||
                        Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["基本情况"].range("B8").value
                        == "厦门百邦税务师事务所有限公司" ||
                        Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["基本情况"].range("B8").value
                        == "厦门明正税务师事务所有限公司" ||
                        Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["基本情况"].range("B8").value
                        == "中汇（厦门）税务师事务所有限公司")
                    {
                        string 版本 =
                            Zifu(Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["首页"].Range["A1"].Value2);
                        if (版本.IndexOf("V2016") >= 0)
                        {
                            if (版本.IndexOf("V20160508") >= 0)
                                MessageBox.Show("该底稿为2016版本，新版本只提供查看功能。");
                            else
                                MessageBox.Show("该底稿非2016最终版本，请使用7.22大暑版对底稿进行升级。");
                            WorkingPaper.版本号 = 2016;
                        }
                        else if (版本.IndexOf("V2017") >= 0)
                        {
                            if (版本.IndexOf("V20171222") < 0)
                                MessageBox.Show("该底稿非最新版本，请升级后使用，以免出现未知错误。");
                            WorkingPaper.版本号 = 2017;
                        }

                        WorkingPaper.Wb = Globals.WPToolAddln.Application.ActiveWorkbook;
                        WorkingPaper.OOO = true;
                        return true;
                    }
                    else
                    {
                        if (Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["基本情况"].range("F11").value
                            == "厦门明正税务师事务所有限公司" ||
                            Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["基本情况"].range("F11").value
                            == "中汇（厦门）税务师事务所有限公司")
                        {
                            string 版本 =
                                Zifu(Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["辅助表"].Range["I1"].Value2);
                            if (版本.IndexOf("V2018") >= 0)
                            {
                                if (版本.IndexOf("V"+WorkingPaper.底稿版本) < 0)
                                    MessageBox.Show("该底稿非最新版本，请升级后使用，以免出现未知错误。");
                                WorkingPaper.版本号 = 2018;

                                WorkingPaper.Wb = Globals.WPToolAddln.Application.ActiveWorkbook;
                                WorkingPaper.OOO = true;
                                return true;
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
            }
            WorkingPaper.OOO = false;
            return false;
        }

        public static void 工作表切换(params string[] str)
        {
            int C;
            if (WorkingPaper.OOO)
            {
                WorkingPaper.Wb.Application.ScreenUpdating = false;
                WorkingPaper.Wb.Worksheets[1].Visible = XlSheetVisibility.xlSheetVisible;
                C = WorkingPaper.Wb.Worksheets.Count;
                for (int i = 2; i <= C; i++)
                {
                    WorkingPaper.Wb.Worksheets[i].Visible = XlSheetVisibility.xlSheetHidden;
                }
                foreach (string ss in str)
                {

                    try
                    {
                        WorkingPaper.Wb.Worksheets[ss].Visible = XlSheetVisibility.xlSheetVisible;

                    }
                    catch (Exception ex)
                    {
                        //Globals.WPToolAddln.Application.ScreenUpdating = true;
                        //System.Windows.Forms.MessageBox.Show("用户操作出现错误：" + ex.Message);
                    }
                }
                WorkingPaper.Wb.Worksheets[1].Visible = XlSheetVisibility.xlSheetVeryHidden;
                WorkingPaper.Wb.Application.ScreenUpdating = true;
            }
        }

        public static double Shuzi(object tar)
        {
            double kk;
            if (tar == null)
                return 0;
            else
                if (double.TryParse(tar.ToString(),out kk))
                return kk;
            return 0;
        }

        public static string Zifu(object tar)
        {
            if (tar == null)
                return "";
            else
                return tar.ToString();
        }

        public static void 自动调整行高(string 表名,string 地址,double width)
        {
            Workbook wrkBook = WorkingPaper.wb打印;
            Worksheet mySheet = wrkBook.Worksheets[表名];
            Range rrng = mySheet.Range[地址];
            Worksheet wrkSheet = wrkBook.Worksheets["调整"];
            wrkSheet.Visible = XlSheetVisibility.xlSheetVisible;
            wrkSheet.Columns[1].WrapText = true;
            wrkSheet.Cells[1, 1].Value = rrng.Value;
            wrkSheet.Columns[1].Font.Size = rrng.Cells[1,1].Font.Size;
            wrkSheet.Columns[1].ColumnWidth = width;
            wrkSheet.Activate();
            wrkSheet.Cells[1, 1].RowHeight = 0;
            wrkSheet.Cells[1, 1].EntireRow.Activate();
            wrkSheet.Cells[1, 1].EntireRow.AutoFit();
            mySheet.Activate();
            rrng.Activate();
            rrng.RowHeight = wrkSheet.Cells[1, 1].RowHeight + 10;
            wrkSheet.Visible = XlSheetVisibility.xlSheetHidden;
        }
    }
}
