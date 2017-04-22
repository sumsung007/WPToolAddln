namespace 百邦所得税汇算底稿工具
{
    partial class WorkingPaper : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public WorkingPaper()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.WPTool = this.Factory.CreateRibbonTab();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.btn新建 = this.Factory.CreateRibbonButton();
            this.tb显示目录 = this.Factory.CreateRibbonToggleButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btn基本情况 = this.Factory.CreateRibbonButton();
            this.btn余额报表 = this.Factory.CreateRibbonButton();
            this.btn税费测算 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btn检查表 = this.Factory.CreateRibbonButton();
            this.btn底稿打印 = this.Factory.CreateRibbonButton();
            this.btn底稿查看 = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.btn客户沟通 = this.Factory.CreateRibbonButton();
            this.btn查看报告 = this.Factory.CreateRibbonButton();
            this.splitButton1 = this.Factory.CreateRibbonSplitButton();
            this.btn导出报告 = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.btn打印报告 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.sb导出数据 = this.Factory.CreateRibbonSplitButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnOUT03 = this.Factory.CreateRibbonButton();
            this.btnOUT07 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.btn工具设置 = this.Factory.CreateRibbonSplitButton();
            this.btn底稿升级 = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.button1 = this.Factory.CreateRibbonButton();
            this.Contact = this.Factory.CreateRibbonSplitButton();
            this.btnUpdata = this.Factory.CreateRibbonButton();
            this.btnGongzhonghao = this.Factory.CreateRibbonButton();
            this.btnHelp = this.Factory.CreateRibbonButton();
            this.btn注册 = this.Factory.CreateRibbonButton();
            this.WPTool.SuspendLayout();
            this.group5.SuspendLayout();
            this.group1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // WPTool
            // 
            this.WPTool.Groups.Add(this.group5);
            this.WPTool.Groups.Add(this.group1);
            this.WPTool.Groups.Add(this.group3);
            this.WPTool.Groups.Add(this.group4);
            this.WPTool.Groups.Add(this.group2);
            this.WPTool.Label = "底稿工具";
            this.WPTool.Name = "WPTool";
            // 
            // group5
            // 
            this.group5.Items.Add(this.btn新建);
            this.group5.Items.Add(this.tb显示目录);
            this.group5.Label = "新建模板";
            this.group5.Name = "group5";
            // 
            // btn新建
            // 
            this.btn新建.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn新建.Image = global::百邦所得税汇算底稿工具.Properties.Resources.ic_content_copy_black_24dp;
            this.btn新建.Label = "新建底稿";
            this.btn新建.Name = "btn新建";
            this.btn新建.ShowImage = true;
            this.btn新建.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn新建_Click);
            // 
            // tb显示目录
            // 
            this.tb显示目录.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.tb显示目录.Image = global::百邦所得税汇算底稿工具.Properties.Resources.ic_list_black_48dp;
            this.tb显示目录.Label = "侧边工具";
            this.tb显示目录.Name = "tb显示目录";
            this.tb显示目录.ShowImage = true;
            this.tb显示目录.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.tb显示目录_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.btn基本情况);
            this.group1.Items.Add(this.btn余额报表);
            this.group1.Items.Add(this.btn税费测算);
            this.group1.Label = "基础资料";
            this.group1.Name = "group1";
            // 
            // btn基本情况
            // 
            this.btn基本情况.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn基本情况.Image = global::百邦所得税汇算底稿工具.Properties.Resources.ic_assignment_ind_black_48dp;
            this.btn基本情况.Label = "基本情况";
            this.btn基本情况.Name = "btn基本情况";
            this.btn基本情况.ShowImage = true;
            this.btn基本情况.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn基本情况_Click);
            // 
            // btn余额报表
            // 
            this.btn余额报表.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn余额报表.Image = global::百邦所得税汇算底稿工具.Properties.Resources.ic_flip_black_36dp;
            this.btn余额报表.Label = "余额报表";
            this.btn余额报表.Name = "btn余额报表";
            this.btn余额报表.ShowImage = true;
            this.btn余额报表.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn余额报表_Click);
            // 
            // btn税费测算
            // 
            this.btn税费测算.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn税费测算.Image = global::百邦所得税汇算底稿工具.Properties.Resources.ic_insert_chart_black_36dp;
            this.btn税费测算.Label = "税费测算";
            this.btn税费测算.Name = "btn税费测算";
            this.btn税费测算.ShowImage = true;
            this.btn税费测算.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn税费测算_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btn检查表);
            this.group3.Items.Add(this.btn底稿打印);
            this.group3.Items.Add(this.btn底稿查看);
            this.group3.Label = "底稿查看";
            this.group3.Name = "group3";
            // 
            // btn检查表
            // 
            this.btn检查表.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn检查表.Image = global::百邦所得税汇算底稿工具.Properties.Resources.ic_search_black_24dp;
            this.btn检查表.Label = "检查表";
            this.btn检查表.Name = "btn检查表";
            this.btn检查表.ShowImage = true;
            this.btn检查表.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn检查表_Click);
            // 
            // btn底稿打印
            // 
            this.btn底稿打印.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn底稿打印.Image = global::百邦所得税汇算底稿工具.Properties.Resources.ic_local_print_shop_black_36dp;
            this.btn底稿打印.Label = "底稿打印";
            this.btn底稿打印.Name = "btn底稿打印";
            this.btn底稿打印.ShowImage = true;
            this.btn底稿打印.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn底稿打印_Click);
            // 
            // btn底稿查看
            // 
            this.btn底稿查看.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn底稿查看.Image = global::百邦所得税汇算底稿工具.Properties.Resources.ic_border_color_black_36dp;
            this.btn底稿查看.Label = "底稿查看";
            this.btn底稿查看.Name = "btn底稿查看";
            this.btn底稿查看.ShowImage = true;
            this.btn底稿查看.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn底稿查看_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.btn客户沟通);
            this.group4.Items.Add(this.btn查看报告);
            this.group4.Items.Add(this.splitButton1);
            this.group4.Label = "报告查看";
            this.group4.Name = "group4";
            // 
            // btn客户沟通
            // 
            this.btn客户沟通.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn客户沟通.Image = global::百邦所得税汇算底稿工具.Properties.Resources.ic_question_answer_black_36dp;
            this.btn客户沟通.Label = "客户沟通";
            this.btn客户沟通.Name = "btn客户沟通";
            this.btn客户沟通.ShowImage = true;
            this.btn客户沟通.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button8_Click);
            // 
            // btn查看报告
            // 
            this.btn查看报告.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn查看报告.Image = global::百邦所得税汇算底稿工具.Properties.Resources.ic_receipt_black_24dp;
            this.btn查看报告.Label = "查看报告";
            this.btn查看报告.Name = "btn查看报告";
            this.btn查看报告.ShowImage = true;
            this.btn查看报告.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn查看报告_Click);
            // 
            // splitButton1
            // 
            this.splitButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton1.Image = global::百邦所得税汇算底稿工具.Properties.Resources.ic_cloud_upload_black_36dp;
            this.splitButton1.Items.Add(this.btn导出报告);
            this.splitButton1.Items.Add(this.separator3);
            this.splitButton1.Items.Add(this.btn打印报告);
            this.splitButton1.Label = "报告导出";
            this.splitButton1.Name = "splitButton1";
            // 
            // btn导出报告
            // 
            this.btn导出报告.Enabled = false;
            this.btn导出报告.Image = global::百邦所得税汇算底稿工具.Properties.Resources.ic_file_upload_black_36dp;
            this.btn导出报告.Label = "输出上传文件";
            this.btn导出报告.Name = "btn导出报告";
            this.btn导出报告.ShowImage = true;
            this.btn导出报告.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn导出报告_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            this.separator3.Title = "  ";
            // 
            // btn打印报告
            // 
            this.btn打印报告.Image = global::百邦所得税汇算底稿工具.Properties.Resources.ic_local_cafe_black_36dp;
            this.btn打印报告.Label = "打印报告";
            this.btn打印报告.Name = "btn打印报告";
            this.btn打印报告.ShowImage = true;
            this.btn打印报告.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn打印报告_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.sb导出数据);
            this.group2.Items.Add(this.btn工具设置);
            this.group2.Items.Add(this.Contact);
            this.group2.Items.Add(this.btn注册);
            this.group2.Label = "V20170422✨彩蛋版✨";
            this.group2.Name = "group2";
            // 
            // sb导出数据
            // 
            this.sb导出数据.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.sb导出数据.Image = global::百邦所得税汇算底稿工具.Properties.Resources.ic_launch_black_36dp;
            this.sb导出数据.Items.Add(this.separator1);
            this.sb导出数据.Items.Add(this.btnOUT03);
            this.sb导出数据.Items.Add(this.btnOUT07);
            this.sb导出数据.Items.Add(this.button2);
            this.sb导出数据.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.sb导出数据.Label = "导出数据";
            this.sb导出数据.Name = "sb导出数据";
            this.sb导出数据.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.splitButton1_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            this.separator1.Title = "¤导出当前可见表格";
            // 
            // btnOUT03
            // 
            this.btnOUT03.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnOUT03.Image = global::百邦所得税汇算底稿工具.Properties.Resources.ms_excel;
            this.btnOUT03.Label = "导出Excel2003文件";
            this.btnOUT03.Name = "btnOUT03";
            this.btnOUT03.ShowImage = true;
            this.btnOUT03.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.导出成03);
            // 
            // btnOUT07
            // 
            this.btnOUT07.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnOUT07.Image = global::百邦所得税汇算底稿工具.Properties.Resources.excel;
            this.btnOUT07.Label = "导出Excel2007文件";
            this.btnOUT07.Name = "btnOUT07";
            this.btnOUT07.ShowImage = true;
            this.btnOUT07.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOUT07_Click);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = global::百邦所得税汇算底稿工具.Properties.Resources.pdf;
            this.button2.Label = "导出为PDF文件";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.导出PDF);
            // 
            // btn工具设置
            // 
            this.btn工具设置.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn工具设置.Image = global::百邦所得税汇算底稿工具.Properties.Resources.ic_settings_applications_black_36dp;
            this.btn工具设置.Items.Add(this.btn底稿升级);
            this.btn工具设置.Items.Add(this.separator2);
            this.btn工具设置.Items.Add(this.button1);
            this.btn工具设置.Label = "高级功能";
            this.btn工具设置.Name = "btn工具设置";
            this.btn工具设置.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn工具设置_Click);
            // 
            // btn底稿升级
            // 
            this.btn底稿升级.Label = "底稿升级（谨慎操作）";
            this.btn底稿升级.Name = "btn底稿升级";
            this.btn底稿升级.ShowImage = true;
            this.btn底稿升级.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.底稿升级_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            this.separator2.Title = "以下为特别修复";
            // 
            // button1
            // 
            this.button1.Label = "修复打印报告权限";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click_1);
            // 
            // Contact
            // 
            this.Contact.ButtonType = Microsoft.Office.Tools.Ribbon.RibbonButtonType.ToggleButton;
            this.Contact.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Contact.Image = global::百邦所得税汇算底稿工具.Properties.Resources.ic_info_outline_black_36dp;
            this.Contact.Items.Add(this.btnUpdata);
            this.Contact.Items.Add(this.btnGongzhonghao);
            this.Contact.Items.Add(this.btnHelp);
            this.Contact.Label = "联系我们";
            this.Contact.Name = "Contact";
            // 
            // btnUpdata
            // 
            this.btnUpdata.Label = "手动检查版本";
            this.btnUpdata.Name = "btnUpdata";
            this.btnUpdata.ShowImage = true;
            this.btnUpdata.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdata_Click);
            // 
            // btnGongzhonghao
            // 
            this.btnGongzhonghao.Label = "关注公众号";
            this.btnGongzhonghao.Name = "btnGongzhonghao";
            this.btnGongzhonghao.ShowImage = true;
            this.btnGongzhonghao.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGongzhonghao_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.Label = "关于 税审底稿工具";
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.ShowImage = true;
            this.btnHelp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHelp_Click);
            // 
            // btn注册
            // 
            this.btn注册.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn注册.Image = global::百邦所得税汇算底稿工具.Properties.Resources.ic_vpn_key_black_36dp;
            this.btn注册.Label = "工具注册";
            this.btn注册.Name = "btn注册";
            this.btn注册.ShowImage = true;
            this.btn注册.Visible = false;
            this.btn注册.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn注册_Click);
            // 
            // WorkingPaper
            // 
            this.Name = "WorkingPaper";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.WPTool);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.WPTool.ResumeLayout(false);
            this.WPTool.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab WPTool;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn基本情况;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn余额报表;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn税费测算;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn检查表;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn底稿打印;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn底稿查看;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn客户沟通;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn查看报告;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn导出报告;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn新建;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton tb显示目录;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn注册;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton sb导出数据;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOUT03;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOUT07;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton btn工具设置;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn底稿升级;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn打印报告;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton Contact;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdata;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGongzhonghao;
    }

    partial class ThisRibbonCollection
    {
        internal WorkingPaper Ribbon1
        {
            get { return this.GetRibbon<WorkingPaper>(); }
        }
    }
}
