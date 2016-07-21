using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Management;

namespace 百邦所得税汇算底稿工具
{
    class CU
    {
        public static Boolean 事项说明()
        {
            if (WorkingPaper.OOO)
            {
                Worksheet SH = WorkingPaper.Wb.Sheets["(二)附表-纳税调整额的审核"];
                SH.Range["A7:E" + SH.Cells[SH.UsedRange.Rows.Count + 1, 1].End[XlDirection.xlUp].Row.ToString()].Value = "";
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
                        Tiaozheng.Add(CU.Shuzi(Nstz[tz, 5]));
                        //Yuanyin.Add("税法规定");
                    }
                }
                if (Xiangmu.Count > 0)
                {
                    string[,] xiangmu = new string[Xiangmu.Count, 1];
                    double[,] jine = new double[Xiangmu.Count, 3];
                    string[,] yuanyin = new string[Xiangmu.Count, 1];
                    int k = 0;
                    foreach (string s in Xiangmu)
                    {
                        xiangmu[k, 0] = s;
                        yuanyin[k, 0] = "税法规定";
                        k++;
                    }
                    k = 0;
                    foreach (double s in Zhangzai)
                    {
                        jine[k, 0] = s;
                        k++;
                    }
                    k = 0;
                    foreach (double s in Shuishou)
                    {
                        jine[k, 1] = s;
                        k++;
                    }
                    k = 0;
                    foreach (double s in Tiaozheng)
                    {
                        jine[k, 2] = s;
                        k++;
                    }
                    SH.Range["A7:A" + (Xiangmu.Count + 6).ToString()].Value2 = xiangmu;
                    SH.Range["B7:D" + (Xiangmu.Count + 6).ToString()].Value2 = jine;
                    SH.Range["E7:E" + (Xiangmu.Count + 6).ToString()].Value2 = yuanyin;
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
                        == "中汇百邦（厦门）税务师事务所有限公司"||
                        Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["基本情况"].range("B8").value
                        == "厦门百邦税务师事务所有限公司" )
                    {
                        if (Zifu(Globals.WPToolAddln.Application.ActiveWorkbook.Worksheets["首页"].Range["A1"].Value2).IndexOf("V20160508") <0)
                            System.Windows.Forms.MessageBox.Show("该底稿非最新版本，请升级后使用，以免出现未知错误。");
                        WorkingPaper.Wb = Globals.WPToolAddln.Application.ActiveWorkbook;
                        WorkingPaper.OOO = true;
                        return true;
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
                try
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
                        WorkingPaper.Wb.Worksheets[ss].Visible = XlSheetVisibility.xlSheetVisible;
                    }
                    WorkingPaper.Wb.Worksheets[1].Visible = XlSheetVisibility.xlSheetVeryHidden;
                    WorkingPaper.Wb.Application.ScreenUpdating = true;
                }
                catch (Exception ex)
                {
                    Globals.WPToolAddln.Application.ScreenUpdating = true;
                    System.Windows.Forms.MessageBox.Show("用户操作出现错误：" + ex.Message);
                }
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
    }
}
