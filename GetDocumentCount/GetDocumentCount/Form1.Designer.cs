namespace GetDocumentCount
{
    partial class Form1
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
            this.ddlFilter = new System.Windows.Forms.ComboBox();
            this.lblFilter = new System.Windows.Forms.Label();
            this.lblURL = new System.Windows.Forms.Label();
            this.txtUrl = new System.Windows.Forms.TextBox();
            this.lblFilterValue = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.pnlCredential = new System.Windows.Forms.Panel();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtUserName = new System.Windows.Forms.TextBox();
            this.lblPassword = new System.Windows.Forms.Label();
            this.lblUserName = new System.Windows.Forms.Label();
            this.pnlToken = new System.Windows.Forms.Panel();
            this.txtClientSecret = new System.Windows.Forms.TextBox();
            this.txtClientID = new System.Windows.Forms.TextBox();
            this.lblClientSecret = new System.Windows.Forms.Label();
            this.lblClientID = new System.Windows.Forms.Label();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtValue = new System.Windows.Forms.TextBox();
            this.lblMsg = new System.Windows.Forms.Label();
            this.pnlCredential.SuspendLayout();
            this.pnlToken.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // ddlFilter
            // 
            this.ddlFilter.FormattingEnabled = true;
            this.ddlFilter.Location = new System.Drawing.Point(194, 318);
            this.ddlFilter.Name = "ddlFilter";
            this.ddlFilter.Size = new System.Drawing.Size(270, 28);
            this.ddlFilter.TabIndex = 0;
            // 
            // lblFilter
            // 
            this.lblFilter.AutoSize = true;
            this.lblFilter.Location = new System.Drawing.Point(94, 321);
            this.lblFilter.Name = "lblFilter";
            this.lblFilter.Size = new System.Drawing.Size(82, 20);
            this.lblFilter.TabIndex = 1;
            this.lblFilter.Text = "Filter Type";
            // 
            // lblURL
            // 
            this.lblURL.AutoSize = true;
            this.lblURL.Location = new System.Drawing.Point(94, 41);
            this.lblURL.Name = "lblURL";
            this.lblURL.Size = new System.Drawing.Size(74, 20);
            this.lblURL.TabIndex = 2;
            this.lblURL.Text = "Site URL";
            // 
            // txtUrl
            // 
            this.txtUrl.Location = new System.Drawing.Point(194, 35);
            this.txtUrl.Name = "txtUrl";
            this.txtUrl.Size = new System.Drawing.Size(675, 26);
            this.txtUrl.TabIndex = 3;
            // 
            // lblFilterValue
            // 
            this.lblFilterValue.AutoSize = true;
            this.lblFilterValue.Location = new System.Drawing.Point(505, 324);
            this.lblFilterValue.Name = "lblFilterValue";
            this.lblFilterValue.Size = new System.Drawing.Size(50, 20);
            this.lblFilterValue.TabIndex = 4;
            this.lblFilterValue.Text = "Value";
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(281, 414);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 29);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(676, 414);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(90, 33);
            this.btnGenerate.TabIndex = 7;
            this.btnGenerate.Text = "Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.BtnGenerate_Click);
            // 
            // pnlCredential
            // 
            this.pnlCredential.Controls.Add(this.txtPassword);
            this.pnlCredential.Controls.Add(this.txtUserName);
            this.pnlCredential.Controls.Add(this.lblPassword);
            this.pnlCredential.Controls.Add(this.lblUserName);
            this.pnlCredential.Location = new System.Drawing.Point(28, 50);
            this.pnlCredential.Name = "pnlCredential";
            this.pnlCredential.Size = new System.Drawing.Size(722, 134);
            this.pnlCredential.TabIndex = 8;
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(167, 81);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '*';
            this.txtPassword.Size = new System.Drawing.Size(502, 26);
            this.txtPassword.TabIndex = 3;
            // 
            // txtUserName
            // 
            this.txtUserName.Location = new System.Drawing.Point(167, 36);
            this.txtUserName.Name = "txtUserName";
            this.txtUserName.Size = new System.Drawing.Size(502, 26);
            this.txtUserName.TabIndex = 2;
            // 
            // lblPassword
            // 
            this.lblPassword.AutoSize = true;
            this.lblPassword.Location = new System.Drawing.Point(35, 88);
            this.lblPassword.Name = "lblPassword";
            this.lblPassword.Size = new System.Drawing.Size(78, 20);
            this.lblPassword.TabIndex = 1;
            this.lblPassword.Text = "Password";
            // 
            // lblUserName
            // 
            this.lblUserName.AutoSize = true;
            this.lblUserName.Location = new System.Drawing.Point(31, 36);
            this.lblUserName.Name = "lblUserName";
            this.lblUserName.Size = new System.Drawing.Size(85, 20);
            this.lblUserName.TabIndex = 0;
            this.lblUserName.Text = "UserName";
            // 
            // pnlToken
            // 
            this.pnlToken.Controls.Add(this.txtClientSecret);
            this.pnlToken.Controls.Add(this.txtClientID);
            this.pnlToken.Controls.Add(this.lblClientSecret);
            this.pnlToken.Controls.Add(this.lblClientID);
            this.pnlToken.Location = new System.Drawing.Point(31, 50);
            this.pnlToken.Name = "pnlToken";
            this.pnlToken.Size = new System.Drawing.Size(719, 134);
            this.pnlToken.TabIndex = 9;
            // 
            // txtClientSecret
            // 
            this.txtClientSecret.Location = new System.Drawing.Point(167, 79);
            this.txtClientSecret.Name = "txtClientSecret";
            this.txtClientSecret.Size = new System.Drawing.Size(502, 26);
            this.txtClientSecret.TabIndex = 3;
            // 
            // txtClientID
            // 
            this.txtClientID.Location = new System.Drawing.Point(167, 30);
            this.txtClientID.Name = "txtClientID";
            this.txtClientID.Size = new System.Drawing.Size(502, 26);
            this.txtClientID.TabIndex = 2;
            // 
            // lblClientSecret
            // 
            this.lblClientSecret.AutoSize = true;
            this.lblClientSecret.Location = new System.Drawing.Point(35, 86);
            this.lblClientSecret.Name = "lblClientSecret";
            this.lblClientSecret.Size = new System.Drawing.Size(100, 20);
            this.lblClientSecret.TabIndex = 1;
            this.lblClientSecret.Text = "Client Secret";
            // 
            // lblClientID
            // 
            this.lblClientID.AutoSize = true;
            this.lblClientID.Location = new System.Drawing.Point(31, 37);
            this.lblClientID.Name = "lblClientID";
            this.lblClientID.Size = new System.Drawing.Size(70, 20);
            this.lblClientID.TabIndex = 0;
            this.lblClientID.Text = "Client ID";
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Location = new System.Drawing.Point(133, 25);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(78, 24);
            this.radioButton1.TabIndex = 10;
            this.radioButton1.Text = "Token";
            this.radioButton1.UseVisualStyleBackColor = true;
            this.radioButton1.CheckedChanged += new System.EventHandler(this.RadioButton1_CheckedChanged);
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Checked = true;
            this.radioButton2.Location = new System.Drawing.Point(392, 24);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(152, 24);
            this.radioButton2.TabIndex = 11;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "User Credentials";
            this.radioButton2.UseVisualStyleBackColor = true;
            this.radioButton2.CheckedChanged += new System.EventHandler(this.RadioButton2_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButton1);
            this.groupBox1.Controls.Add(this.radioButton2);
            this.groupBox1.Controls.Add(this.pnlCredential);
            this.groupBox1.Controls.Add(this.pnlToken);
            this.groupBox1.Location = new System.Drawing.Point(86, 76);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(783, 207);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Connect Using";
            // 
            // txtValue
            // 
            this.txtValue.Location = new System.Drawing.Point(561, 321);
            this.txtValue.Name = "txtValue";
            this.txtValue.Size = new System.Drawing.Size(308, 26);
            this.txtValue.TabIndex = 13;
            // 
            // lblMsg
            // 
            this.lblMsg.AutoSize = true;
            this.lblMsg.Location = new System.Drawing.Point(98, 375);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(0, 20);
            this.lblMsg.TabIndex = 14;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(986, 506);
            this.Controls.Add(this.lblMsg);
            this.Controls.Add(this.txtValue);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.lblFilterValue);
            this.Controls.Add(this.txtUrl);
            this.Controls.Add(this.lblURL);
            this.Controls.Add(this.lblFilter);
            this.Controls.Add(this.ddlFilter);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.pnlCredential.ResumeLayout(false);
            this.pnlCredential.PerformLayout();
            this.pnlToken.ResumeLayout(false);
            this.pnlToken.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox ddlFilter;
        private System.Windows.Forms.Label lblFilter;
        private System.Windows.Forms.Label lblURL;
        private System.Windows.Forms.TextBox txtUrl;
        private System.Windows.Forms.Label lblFilterValue;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.Panel pnlCredential;
        private System.Windows.Forms.Panel pnlToken;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.Label lblUserName;
        private System.Windows.Forms.Label lblClientID;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.TextBox txtUserName;
        private System.Windows.Forms.Label lblPassword;
        private System.Windows.Forms.Label lblClientSecret;
        private System.Windows.Forms.TextBox txtClientSecret;
        private System.Windows.Forms.TextBox txtClientID;
        private System.Windows.Forms.TextBox txtValue;
        private System.Windows.Forms.Label lblMsg;
    }
}

