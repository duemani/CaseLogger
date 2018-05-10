namespace CaseLogger
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
            this.acgmeWebBrowser = new System.Windows.Forms.WebBrowser();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.openExcelStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.convertXPSToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.getCPTsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.logonToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.submitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.nameToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.yearToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem4 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem5 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem6 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem7 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem8 = new System.Windows.Forms.ToolStripMenuItem();
            this.residentRoleToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.seniorToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.leadToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.assistantToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openXPSFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.sENMODEToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.senmodelabel = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // acgmeWebBrowser
            // 
            this.acgmeWebBrowser.Dock = System.Windows.Forms.DockStyle.Fill;
            this.acgmeWebBrowser.Location = new System.Drawing.Point(0, 24);
            this.acgmeWebBrowser.MinimumSize = new System.Drawing.Size(20, 20);
            this.acgmeWebBrowser.Name = "acgmeWebBrowser";
            this.acgmeWebBrowser.Size = new System.Drawing.Size(1071, 609);
            this.acgmeWebBrowser.TabIndex = 0;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openExcelStripMenuItem,
            this.convertXPSToolStripMenuItem,
            this.getCPTsToolStripMenuItem,
            this.logonToolStripMenuItem,
            this.submitToolStripMenuItem,
            this.aboutToolStripMenuItem,
            this.settingsToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1071, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // openExcelStripMenuItem
            // 
            this.openExcelStripMenuItem.Name = "openExcelStripMenuItem";
            this.openExcelStripMenuItem.Size = new System.Drawing.Size(77, 20);
            this.openExcelStripMenuItem.Text = "Open Excel";
            this.openExcelStripMenuItem.Click += new System.EventHandler(this.openExcelStripMenuItem_Click_1);
            // 
            // convertXPSToolStripMenuItem
            // 
            this.convertXPSToolStripMenuItem.Name = "convertXPSToolStripMenuItem";
            this.convertXPSToolStripMenuItem.Size = new System.Drawing.Size(84, 20);
            this.convertXPSToolStripMenuItem.Text = "Convert XPS";
            this.convertXPSToolStripMenuItem.Click += new System.EventHandler(this.convertXPSToolStripMenuItem_Click);
            // 
            // getCPTsToolStripMenuItem
            // 
            this.getCPTsToolStripMenuItem.Name = "getCPTsToolStripMenuItem";
            this.getCPTsToolStripMenuItem.Size = new System.Drawing.Size(66, 20);
            this.getCPTsToolStripMenuItem.Text = "Get CPTs";
            this.getCPTsToolStripMenuItem.Click += new System.EventHandler(this.getCPTsToolStripMenuItem_Click);
            // 
            // logonToolStripMenuItem
            // 
            this.logonToolStripMenuItem.Name = "logonToolStripMenuItem";
            this.logonToolStripMenuItem.Size = new System.Drawing.Size(53, 20);
            this.logonToolStripMenuItem.Text = "Logon";
            this.logonToolStripMenuItem.Click += new System.EventHandler(this.logonToolStripMenuItem_Click);
            // 
            // submitToolStripMenuItem
            // 
            this.submitToolStripMenuItem.Name = "submitToolStripMenuItem";
            this.submitToolStripMenuItem.Size = new System.Drawing.Size(43, 20);
            this.submitToolStripMenuItem.Text = "Next";
            this.submitToolStripMenuItem.Click += new System.EventHandler(this.submitToolStripMenuItem_Click);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(52, 20);
            this.aboutToolStripMenuItem.Text = "About";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // settingsToolStripMenuItem
            // 
            this.settingsToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.settingsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.nameToolStripMenuItem,
            this.yearToolStripMenuItem,
            this.residentRoleToolStripMenuItem,
            this.sENMODEToolStripMenuItem});
            this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            this.settingsToolStripMenuItem.Size = new System.Drawing.Size(61, 20);
            this.settingsToolStripMenuItem.Text = "Settings";
            // 
            // nameToolStripMenuItem
            // 
            this.nameToolStripMenuItem.Name = "nameToolStripMenuItem";
            this.nameToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.nameToolStripMenuItem.Text = "Name";
            this.nameToolStripMenuItem.Click += new System.EventHandler(this.nameToolStripMenuItem_Click);
            // 
            // yearToolStripMenuItem
            // 
            this.yearToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem2,
            this.toolStripMenuItem3,
            this.toolStripMenuItem4,
            this.toolStripMenuItem5,
            this.toolStripMenuItem6,
            this.toolStripMenuItem7,
            this.toolStripMenuItem8});
            this.yearToolStripMenuItem.Name = "yearToolStripMenuItem";
            this.yearToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.yearToolStripMenuItem.Text = "Year";
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(80, 22);
            this.toolStripMenuItem2.Text = "1";
            this.toolStripMenuItem2.Click += new System.EventHandler(this.toolStripMenuItem2_Click);
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(80, 22);
            this.toolStripMenuItem3.Text = "2";
            this.toolStripMenuItem3.Click += new System.EventHandler(this.toolStripMenuItem3_Click);
            // 
            // toolStripMenuItem4
            // 
            this.toolStripMenuItem4.Name = "toolStripMenuItem4";
            this.toolStripMenuItem4.Size = new System.Drawing.Size(80, 22);
            this.toolStripMenuItem4.Text = "3";
            this.toolStripMenuItem4.Click += new System.EventHandler(this.toolStripMenuItem4_Click);
            // 
            // toolStripMenuItem5
            // 
            this.toolStripMenuItem5.Name = "toolStripMenuItem5";
            this.toolStripMenuItem5.Size = new System.Drawing.Size(80, 22);
            this.toolStripMenuItem5.Text = "4";
            this.toolStripMenuItem5.Click += new System.EventHandler(this.toolStripMenuItem5_Click);
            // 
            // toolStripMenuItem6
            // 
            this.toolStripMenuItem6.Name = "toolStripMenuItem6";
            this.toolStripMenuItem6.Size = new System.Drawing.Size(80, 22);
            this.toolStripMenuItem6.Text = "5";
            this.toolStripMenuItem6.Click += new System.EventHandler(this.toolStripMenuItem6_Click);
            // 
            // toolStripMenuItem7
            // 
            this.toolStripMenuItem7.Name = "toolStripMenuItem7";
            this.toolStripMenuItem7.Size = new System.Drawing.Size(80, 22);
            this.toolStripMenuItem7.Text = "6";
            this.toolStripMenuItem7.Click += new System.EventHandler(this.toolStripMenuItem7_Click);
            // 
            // toolStripMenuItem8
            // 
            this.toolStripMenuItem8.Name = "toolStripMenuItem8";
            this.toolStripMenuItem8.Size = new System.Drawing.Size(80, 22);
            this.toolStripMenuItem8.Text = "7";
            this.toolStripMenuItem8.Click += new System.EventHandler(this.toolStripMenuItem8_Click);
            // 
            // residentRoleToolStripMenuItem
            // 
            this.residentRoleToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.seniorToolStripMenuItem,
            this.leadToolStripMenuItem,
            this.assistantToolStripMenuItem});
            this.residentRoleToolStripMenuItem.Name = "residentRoleToolStripMenuItem";
            this.residentRoleToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.residentRoleToolStripMenuItem.Text = "Resident Role";
            // 
            // seniorToolStripMenuItem
            // 
            this.seniorToolStripMenuItem.Name = "seniorToolStripMenuItem";
            this.seniorToolStripMenuItem.Size = new System.Drawing.Size(121, 22);
            this.seniorToolStripMenuItem.Text = "Senior";
            this.seniorToolStripMenuItem.Click += new System.EventHandler(this.seniorToolStripMenuItem_Click);
            // 
            // leadToolStripMenuItem
            // 
            this.leadToolStripMenuItem.Name = "leadToolStripMenuItem";
            this.leadToolStripMenuItem.Size = new System.Drawing.Size(121, 22);
            this.leadToolStripMenuItem.Text = "Lead";
            this.leadToolStripMenuItem.Click += new System.EventHandler(this.leadToolStripMenuItem_Click);
            // 
            // assistantToolStripMenuItem
            // 
            this.assistantToolStripMenuItem.Name = "assistantToolStripMenuItem";
            this.assistantToolStripMenuItem.Size = new System.Drawing.Size(121, 22);
            this.assistantToolStripMenuItem.Text = "Assistant";
            this.assistantToolStripMenuItem.Click += new System.EventHandler(this.assistantToolStripMenuItem_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(579, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "label1";
            this.label1.Visible = false;
            // 
            // sENMODEToolStripMenuItem
            // 
            this.sENMODEToolStripMenuItem.Name = "sENMODEToolStripMenuItem";
            this.sENMODEToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.sENMODEToolStripMenuItem.Text = "SEN MODE";
            this.sENMODEToolStripMenuItem.Click += new System.EventHandler(this.sENMODEToolStripMenuItem_Click);
            // 
            // senmodelabel
            // 
            this.senmodelabel.AutoSize = true;
            this.senmodelabel.ForeColor = System.Drawing.Color.Red;
            this.senmodelabel.Location = new System.Drawing.Point(847, 5);
            this.senmodelabel.Name = "senmodelabel";
            this.senmodelabel.Size = new System.Drawing.Size(64, 13);
            this.senmodelabel.TabIndex = 3;
            this.senmodelabel.Text = "SEN MODE";
            this.senmodelabel.Visible = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1071, 633);
            this.Controls.Add(this.senmodelabel);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.acgmeWebBrowser);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Case Logger";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.WebBrowser acgmeWebBrowser;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem logonToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openExcelStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem getCPTsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem submitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem convertXPSToolStripMenuItem;
        private System.Windows.Forms.OpenFileDialog openXPSFileDialog;
        private System.Windows.Forms.ToolStripMenuItem settingsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem nameToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem yearToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem3;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem4;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem5;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem6;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem7;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem8;
        private System.Windows.Forms.ToolStripMenuItem residentRoleToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem seniorToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem leadToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem assistantToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ToolStripMenuItem sENMODEToolStripMenuItem;
        private System.Windows.Forms.Label senmodelabel;
    }
}

