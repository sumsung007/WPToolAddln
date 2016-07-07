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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.lv待选 = new System.Windows.Forms.ListView();
            this.lv选中 = new System.Windows.Forms.ListView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.btn识别 = new System.Windows.Forms.Button();
            this.btn全移 = new System.Windows.Forms.Button();
            this.btn移出 = new System.Windows.Forms.Button();
            this.btn全选 = new System.Windows.Forms.Button();
            this.btn选中 = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btn取消 = new System.Windows.Forms.Button();
            this.btn打印 = new System.Windows.Forms.Button();
            this.tableLayoutPanel1.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 119F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.lv待选, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.lv选中, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.panel1, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.panel2, 2, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 49F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(903, 401);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // lv待选
            // 
            this.lv待选.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lv待选.Location = new System.Drawing.Point(3, 3);
            this.lv待选.Name = "lv待选";
            this.lv待选.Size = new System.Drawing.Size(386, 346);
            this.lv待选.TabIndex = 0;
            this.lv待选.UseCompatibleStateImageBehavior = false;
            this.lv待选.View = System.Windows.Forms.View.Details;
            // 
            // lv选中
            // 
            this.lv选中.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lv选中.Location = new System.Drawing.Point(514, 3);
            this.lv选中.Name = "lv选中";
            this.lv选中.Size = new System.Drawing.Size(386, 346);
            this.lv选中.TabIndex = 2;
            this.lv选中.UseCompatibleStateImageBehavior = false;
            this.lv选中.View = System.Windows.Forms.View.Details;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Controls.Add(this.btn识别);
            this.panel1.Controls.Add(this.btn全移);
            this.panel1.Controls.Add(this.btn移出);
            this.panel1.Controls.Add(this.btn全选);
            this.panel1.Controls.Add(this.btn选中);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(395, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(113, 346);
            this.panel1.TabIndex = 3;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(8, 9);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(100, 50);
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // btn识别
            // 
            this.btn识别.Location = new System.Drawing.Point(8, 252);
            this.btn识别.Name = "btn识别";
            this.btn识别.Size = new System.Drawing.Size(93, 23);
            this.btn识别.TabIndex = 0;
            this.btn识别.Text = "智能识别";
            this.btn识别.UseVisualStyleBackColor = true;
            // 
            // btn全移
            // 
            this.btn全移.Location = new System.Drawing.Point(8, 172);
            this.btn全移.Name = "btn全移";
            this.btn全移.Size = new System.Drawing.Size(93, 23);
            this.btn全移.TabIndex = 0;
            this.btn全移.Text = "<-移出全部";
            this.btn全移.UseVisualStyleBackColor = true;
            // 
            // btn移出
            // 
            this.btn移出.Location = new System.Drawing.Point(8, 143);
            this.btn移出.Name = "btn移出";
            this.btn移出.Size = new System.Drawing.Size(93, 23);
            this.btn移出.TabIndex = 0;
            this.btn移出.Text = "<--移出";
            this.btn移出.UseVisualStyleBackColor = true;
            // 
            // btn全选
            // 
            this.btn全选.Location = new System.Drawing.Point(8, 104);
            this.btn全选.Name = "btn全选";
            this.btn全选.Size = new System.Drawing.Size(93, 23);
            this.btn全选.TabIndex = 0;
            this.btn全选.Text = "选中全部->";
            this.btn全选.UseVisualStyleBackColor = true;
            // 
            // btn选中
            // 
            this.btn选中.Location = new System.Drawing.Point(8, 75);
            this.btn选中.Name = "btn选中";
            this.btn选中.Size = new System.Drawing.Size(93, 23);
            this.btn选中.TabIndex = 0;
            this.btn选中.Text = "选中-->";
            this.btn选中.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.btn取消);
            this.panel2.Controls.Add(this.btn打印);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(514, 355);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(386, 43);
            this.panel2.TabIndex = 4;
            // 
            // btn取消
            // 
            this.btn取消.Location = new System.Drawing.Point(251, 11);
            this.btn取消.Name = "btn取消";
            this.btn取消.Size = new System.Drawing.Size(75, 23);
            this.btn取消.TabIndex = 0;
            this.btn取消.Text = "取消打印";
            this.btn取消.UseVisualStyleBackColor = true;
            // 
            // btn打印
            // 
            this.btn打印.Location = new System.Drawing.Point(21, 11);
            this.btn打印.Name = "btn打印";
            this.btn打印.Size = new System.Drawing.Size(75, 23);
            this.btn打印.TabIndex = 0;
            this.btn打印.Text = "批量打印";
            this.btn打印.UseVisualStyleBackColor = true;
            this.btn打印.Click += new System.EventHandler(this.btn打印_Click);
            // 
            // 底稿打印
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(903, 401);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "底稿打印";
            this.Text = "Form1";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.ListView lv待选;
        private System.Windows.Forms.ListView lv选中;
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
    }
}