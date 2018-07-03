namespace 百邦所得税汇算底稿工具
{
    partial class 底稿打印
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.ListViewGroup listViewGroup1 = new System.Windows.Forms.ListViewGroup("必须打印", System.Windows.Forms.HorizontalAlignment.Center);
            System.Windows.Forms.ListViewGroup listViewGroup2 = new System.Windows.Forms.ListViewGroup("选择打印", System.Windows.Forms.HorizontalAlignment.Center);
            System.Windows.Forms.ListViewGroup listViewGroup3 = new System.Windows.Forms.ListViewGroup("无需打印", System.Windows.Forms.HorizontalAlignment.Center);
            System.Windows.Forms.ListViewItem listViewItem1 = new System.Windows.Forms.ListViewItem(new string[] {
            "333",
            "有数"}, -1);
            System.Windows.Forms.ListViewItem listViewItem2 = new System.Windows.Forms.ListViewItem("xuanzeda");
            System.Windows.Forms.ListViewItem listViewItem3 = new System.Windows.Forms.ListViewItem("无需打印");
            System.Windows.Forms.ListViewGroup listViewGroup4 = new System.Windows.Forms.ListViewGroup("必须打印", System.Windows.Forms.HorizontalAlignment.Center);
            System.Windows.Forms.ListViewGroup listViewGroup5 = new System.Windows.Forms.ListViewGroup("选择打印", System.Windows.Forms.HorizontalAlignment.Center);
            System.Windows.Forms.ListViewGroup listViewGroup6 = new System.Windows.Forms.ListViewGroup("无需打印", System.Windows.Forms.HorizontalAlignment.Center);
            System.Windows.Forms.ListViewItem listViewItem4 = new System.Windows.Forms.ListViewItem(new string[] {
            "333",
            "有数"}, -1);
            System.Windows.Forms.ListViewItem listViewItem5 = new System.Windows.Forms.ListViewItem("xuanzeda");
            System.Windows.Forms.ListViewItem listViewItem6 = new System.Windows.Forms.ListViewItem("无需打印");
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(底稿打印));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.lv待选 = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btn全移 = new System.Windows.Forms.Button();
            this.btn移出 = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.btn识别 = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btn全选 = new System.Windows.Forms.Button();
            this.btn选中 = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btn取消 = new System.Windows.Forms.Button();
            this.btn打印 = new System.Windows.Forms.Button();
            this.lv选中 = new System.Windows.Forms.ListView();
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader4 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.label1 = new System.Windows.Forms.Label();
            this.tableLayoutPanel1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 98F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.lv待选, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.panel1, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.panel2, 2, 1);
            this.tableLayoutPanel1.Controls.Add(this.lv选中, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.label1, 0, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 39F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1067, 448);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // lv待选
            // 
            this.lv待选.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2});
            this.lv待选.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lv待选.FullRowSelect = true;
            listViewGroup1.Header = "必须打印";
            listViewGroup1.HeaderAlignment = System.Windows.Forms.HorizontalAlignment.Center;
            listViewGroup1.Name = "MustGroup";
            listViewGroup2.Header = "选择打印";
            listViewGroup2.HeaderAlignment = System.Windows.Forms.HorizontalAlignment.Center;
            listViewGroup2.Name = "ChooseGroup";
            listViewGroup3.Header = "无需打印";
            listViewGroup3.HeaderAlignment = System.Windows.Forms.HorizontalAlignment.Center;
            listViewGroup3.Name = "NonGroup";
            this.lv待选.Groups.AddRange(new System.Windows.Forms.ListViewGroup[] {
            listViewGroup1,
            listViewGroup2,
            listViewGroup3});
            this.lv待选.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            listViewItem1.Group = listViewGroup1;
            listViewItem2.Group = listViewGroup2;
            listViewItem3.Group = listViewGroup3;
            this.lv待选.Items.AddRange(new System.Windows.Forms.ListViewItem[] {
            listViewItem1,
            listViewItem2,
            listViewItem3});
            this.lv待选.Location = new System.Drawing.Point(2, 2);
            this.lv待选.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.lv待选.Name = "lv待选";
            this.lv待选.Size = new System.Drawing.Size(480, 405);
            this.lv待选.TabIndex = 0;
            this.lv待选.UseCompatibleStateImageBehavior = false;
            this.lv待选.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "表名";
            this.columnHeader1.Width = 377;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "有效性";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Controls.Add(this.btn识别);
            this.panel1.Controls.Add(this.groupBox2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(486, 2);
            this.panel1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(94, 405);
            this.panel1.TabIndex = 3;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btn全移);
            this.groupBox1.Controls.Add(this.btn移出);
            this.groupBox1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.groupBox1.Location = new System.Drawing.Point(2, 173);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Size = new System.Drawing.Size(88, 97);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "移出";
            // 
            // btn全移
            // 
            this.btn全移.Location = new System.Drawing.Point(7, 58);
            this.btn全移.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btn全移.Name = "btn全移";
            this.btn全移.Size = new System.Drawing.Size(77, 27);
            this.btn全移.TabIndex = 0;
            this.btn全移.Text = "<-移出全部";
            this.btn全移.UseVisualStyleBackColor = true;
            this.btn全移.Click += new System.EventHandler(this.btn全移_Click);
            // 
            // btn移出
            // 
            this.btn移出.Location = new System.Drawing.Point(7, 19);
            this.btn移出.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btn移出.Name = "btn移出";
            this.btn移出.Size = new System.Drawing.Size(77, 27);
            this.btn移出.TabIndex = 0;
            this.btn移出.Text = "<--移出";
            this.btn移出.UseVisualStyleBackColor = true;
            this.btn移出.Click += new System.EventHandler(this.btn移出_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(0, 0);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(93, 52);
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // btn识别
            // 
            this.btn识别.Location = new System.Drawing.Point(9, 283);
            this.btn识别.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btn识别.Name = "btn识别";
            this.btn识别.Size = new System.Drawing.Size(77, 27);
            this.btn识别.TabIndex = 0;
            this.btn识别.Text = "智能识别";
            this.btn识别.UseVisualStyleBackColor = true;
            this.btn识别.Click += new System.EventHandler(this.btn识别_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btn全选);
            this.groupBox2.Controls.Add(this.btn选中);
            this.groupBox2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.groupBox2.Location = new System.Drawing.Point(2, 57);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox2.Size = new System.Drawing.Size(88, 97);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "选中";
            // 
            // btn全选
            // 
            this.btn全选.Location = new System.Drawing.Point(6, 58);
            this.btn全选.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btn全选.Name = "btn全选";
            this.btn全选.Size = new System.Drawing.Size(77, 27);
            this.btn全选.TabIndex = 0;
            this.btn全选.Text = "选中全部->";
            this.btn全选.UseVisualStyleBackColor = true;
            this.btn全选.Click += new System.EventHandler(this.btn全选_Click);
            // 
            // btn选中
            // 
            this.btn选中.Location = new System.Drawing.Point(7, 19);
            this.btn选中.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btn选中.Name = "btn选中";
            this.btn选中.Size = new System.Drawing.Size(77, 27);
            this.btn选中.TabIndex = 0;
            this.btn选中.Text = "选中-->";
            this.btn选中.UseVisualStyleBackColor = true;
            this.btn选中.Click += new System.EventHandler(this.btn选中_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.btn取消);
            this.panel2.Controls.Add(this.btn打印);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(584, 411);
            this.panel2.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(481, 35);
            this.panel2.TabIndex = 4;
            // 
            // btn取消
            // 
            this.btn取消.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btn取消.Location = new System.Drawing.Point(256, 2);
            this.btn取消.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btn取消.Name = "btn取消";
            this.btn取消.Size = new System.Drawing.Size(77, 27);
            this.btn取消.TabIndex = 0;
            this.btn取消.Text = "取消打印";
            this.btn取消.UseVisualStyleBackColor = true;
            this.btn取消.Click += new System.EventHandler(this.btn取消_Click);
            // 
            // btn打印
            // 
            this.btn打印.ForeColor = System.Drawing.Color.Green;
            this.btn打印.Location = new System.Drawing.Point(10, 2);
            this.btn打印.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btn打印.Name = "btn打印";
            this.btn打印.Size = new System.Drawing.Size(77, 27);
            this.btn打印.TabIndex = 0;
            this.btn打印.Text = "批量打印";
            this.btn打印.UseVisualStyleBackColor = true;
            this.btn打印.Click += new System.EventHandler(this.btn打印_Click);
            // 
            // lv选中
            // 
            this.lv选中.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader3,
            this.columnHeader4});
            this.lv选中.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lv选中.FullRowSelect = true;
            listViewGroup4.Header = "必须打印";
            listViewGroup4.HeaderAlignment = System.Windows.Forms.HorizontalAlignment.Center;
            listViewGroup4.Name = "MustGroup";
            listViewGroup5.Header = "选择打印";
            listViewGroup5.HeaderAlignment = System.Windows.Forms.HorizontalAlignment.Center;
            listViewGroup5.Name = "ChooseGroup";
            listViewGroup6.Header = "无需打印";
            listViewGroup6.HeaderAlignment = System.Windows.Forms.HorizontalAlignment.Center;
            listViewGroup6.Name = "NonGroup";
            this.lv选中.Groups.AddRange(new System.Windows.Forms.ListViewGroup[] {
            listViewGroup4,
            listViewGroup5,
            listViewGroup6});
            this.lv选中.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            listViewItem4.Group = listViewGroup4;
            listViewItem5.Group = listViewGroup5;
            listViewItem6.Group = listViewGroup6;
            this.lv选中.Items.AddRange(new System.Windows.Forms.ListViewItem[] {
            listViewItem4,
            listViewItem5,
            listViewItem6});
            this.lv选中.Location = new System.Drawing.Point(584, 2);
            this.lv选中.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.lv选中.Name = "lv选中";
            this.lv选中.Size = new System.Drawing.Size(481, 405);
            this.lv选中.TabIndex = 5;
            this.lv选中.UseCompatibleStateImageBehavior = false;
            this.lv选中.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "表名";
            this.columnHeader3.Width = 380;
            // 
            // columnHeader4
            // 
            this.columnHeader4.Text = "有效性";
            // 
            // label1
            // 
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(2, 409);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(480, 39);
            this.label1.TabIndex = 6;
            this.label1.Text = "STEP.1  选择需要打印的工作表";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // 底稿打印
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1067, 448);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "底稿打印";
            this.Text = "♪(^∇^*)  底稿打印";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }
               

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.ListView lv待选;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btn识别;
        private System.Windows.Forms.Button btn全移;
        private System.Windows.Forms.Button btn移出;
        private System.Windows.Forms.Button btn全选;
        private System.Windows.Forms.Button btn选中;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button btn取消;
        private System.Windows.Forms.Button btn打印;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.ListView lv选中;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.ColumnHeader columnHeader4;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label1;
    }
}