namespace AutoInputOnsite
{
    partial class Menu
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Menu));
            this.cbWeb = new System.Windows.Forms.CheckBox();
            this.cbInsatu = new System.Windows.Forms.CheckBox();
            this.Label1 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.Timer1 = new System.Windows.Forms.Timer(this.components);
            this.pb2 = new System.Windows.Forms.ProgressBar();
            this.pb1 = new System.Windows.Forms.ProgressBar();
            this.cbHyouji = new System.Windows.Forms.CheckBox();
            this.btnAll = new System.Windows.Forms.Button();
            this.btnNouhin = new System.Windows.Forms.Button();
            this.btnMs = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // cbWeb
            // 
            this.cbWeb.AutoSize = true;
            this.cbWeb.Checked = true;
            this.cbWeb.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbWeb.Font = new System.Drawing.Font("HGP創英ﾌﾟﾚｾﾞﾝｽEB", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cbWeb.ForeColor = System.Drawing.Color.GhostWhite;
            this.cbWeb.Location = new System.Drawing.Point(90, 117);
            this.cbWeb.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.cbWeb.Name = "cbWeb";
            this.cbWeb.Size = new System.Drawing.Size(81, 19);
            this.cbWeb.TabIndex = 18;
            this.cbWeb.Text = "Web表示";
            this.cbWeb.UseVisualStyleBackColor = true;
            // 
            // cbInsatu
            // 
            this.cbInsatu.AutoSize = true;
            this.cbInsatu.Checked = true;
            this.cbInsatu.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbInsatu.Font = new System.Drawing.Font("HGP創英ﾌﾟﾚｾﾞﾝｽEB", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cbInsatu.ForeColor = System.Drawing.Color.GhostWhite;
            this.cbInsatu.Location = new System.Drawing.Point(15, 117);
            this.cbInsatu.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.cbInsatu.Name = "cbInsatu";
            this.cbInsatu.Size = new System.Drawing.Size(56, 19);
            this.cbInsatu.TabIndex = 17;
            this.cbInsatu.Text = "印刷";
            this.cbInsatu.UseVisualStyleBackColor = true;
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.BackColor = System.Drawing.Color.Transparent;
            this.Label1.Font = new System.Drawing.Font("HGP創英ﾌﾟﾚｾﾞﾝｽEB", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label1.ForeColor = System.Drawing.Color.GhostWhite;
            this.Label1.Location = new System.Drawing.Point(12, 18);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(112, 15);
            this.Label1.TabIndex = 15;
            this.Label1.Text = "見出＆明細作成";
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.BackColor = System.Drawing.Color.Transparent;
            this.Label2.Font = new System.Drawing.Font("HGP創英ﾌﾟﾚｾﾞﾝｽEB", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label2.ForeColor = System.Drawing.Color.GhostWhite;
            this.Label2.Location = new System.Drawing.Point(12, 52);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(97, 15);
            this.Label2.TabIndex = 16;
            this.Label2.Text = "納期指定発注";
            // 
            // Timer1
            // 
            this.Timer1.Interval = 10;
            this.Timer1.Tick += new System.EventHandler(this.Timer1_Tick_1);
            // 
            // pb2
            // 
            this.pb2.Location = new System.Drawing.Point(130, 41);
            this.pb2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.pb2.Name = "pb2";
            this.pb2.Size = new System.Drawing.Size(506, 26);
            this.pb2.TabIndex = 14;
            // 
            // pb1
            // 
            this.pb1.Location = new System.Drawing.Point(130, 7);
            this.pb1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.pb1.Name = "pb1";
            this.pb1.Size = new System.Drawing.Size(506, 26);
            this.pb1.TabIndex = 13;
            // 
            // cbHyouji
            // 
            this.cbHyouji.AutoSize = true;
            this.cbHyouji.Font = new System.Drawing.Font("HGP創英ﾌﾟﾚｾﾞﾝｽEB", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cbHyouji.ForeColor = System.Drawing.Color.GhostWhite;
            this.cbHyouji.Location = new System.Drawing.Point(550, 117);
            this.cbHyouji.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.cbHyouji.Name = "cbHyouji";
            this.cbHyouji.Size = new System.Drawing.Size(86, 19);
            this.cbHyouji.TabIndex = 12;
            this.cbHyouji.Text = "明細表示";
            this.cbHyouji.UseVisualStyleBackColor = true;
            this.cbHyouji.Visible = false;
            // 
            // btnAll
            // 
            this.btnAll.Font = new System.Drawing.Font("HGP創英ﾌﾟﾚｾﾞﾝｽEB", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnAll.ForeColor = System.Drawing.Color.Blue;
            this.btnAll.Location = new System.Drawing.Point(385, 75);
            this.btnAll.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnAll.Name = "btnAll";
            this.btnAll.Size = new System.Drawing.Size(251, 34);
            this.btnAll.TabIndex = 11;
            this.btnAll.Text = "見出＆明細作成 And 納期指定発注";
            this.btnAll.UseVisualStyleBackColor = true;
            this.btnAll.Click += new System.EventHandler(this.btnAll_Click_1);
            // 
            // btnNouhin
            // 
            this.btnNouhin.Font = new System.Drawing.Font("HGP創英ﾌﾟﾚｾﾞﾝｽEB", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnNouhin.ForeColor = System.Drawing.Color.Blue;
            this.btnNouhin.Location = new System.Drawing.Point(140, 75);
            this.btnNouhin.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnNouhin.Name = "btnNouhin";
            this.btnNouhin.Size = new System.Drawing.Size(118, 34);
            this.btnNouhin.TabIndex = 10;
            this.btnNouhin.Text = "納期指定発注";
            this.btnNouhin.UseVisualStyleBackColor = true;
            this.btnNouhin.Click += new System.EventHandler(this.btnNouhin_Click);
            // 
            // btnMs
            // 
            this.btnMs.Font = new System.Drawing.Font("HGP創英ﾌﾟﾚｾﾞﾝｽEB", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnMs.ForeColor = System.Drawing.Color.Blue;
            this.btnMs.Location = new System.Drawing.Point(12, 75);
            this.btnMs.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnMs.Name = "btnMs";
            this.btnMs.Size = new System.Drawing.Size(122, 34);
            this.btnMs.TabIndex = 9;
            this.btnMs.Text = "見出＆明細作成";
            this.btnMs.UseVisualStyleBackColor = true;
            this.btnMs.Click += new System.EventHandler(this.btnMs_Click_1);
            // 
            // Menu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.ClientSize = new System.Drawing.Size(648, 147);
            this.Controls.Add(this.cbWeb);
            this.Controls.Add(this.cbInsatu);
            this.Controls.Add(this.Label1);
            this.Controls.Add(this.Label2);
            this.Controls.Add(this.pb2);
            this.Controls.Add(this.pb1);
            this.Controls.Add(this.cbHyouji);
            this.Controls.Add(this.btnAll);
            this.Controls.Add(this.btnNouhin);
            this.Controls.Add(this.btnMs);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Menu";
            this.Text = "Onsite 自動化 （見出＆明細作成 And 納期指定発注）";
            this.Load += new System.EventHandler(this.Menu_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.CheckBox cbWeb;
        internal System.Windows.Forms.CheckBox cbInsatu;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.Timer Timer1;
        internal System.Windows.Forms.ProgressBar pb2;
        internal System.Windows.Forms.ProgressBar pb1;
        internal System.Windows.Forms.CheckBox cbHyouji;
        internal System.Windows.Forms.Button btnAll;
        internal System.Windows.Forms.Button btnNouhin;
        internal System.Windows.Forms.Button btnMs;

    }
}