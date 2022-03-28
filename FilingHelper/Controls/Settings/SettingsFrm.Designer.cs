namespace FilingHelper.Controls.Settings
{
    partial class SettingsFrm
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
            this.btnSaveSetting = new System.Windows.Forms.Button();
            this.ctrlMenu = new HelperUtils.SideMenu();
            this.btnCancel = new System.Windows.Forms.Button();
            this.ctrlAttachments = new FilingHelper.Controls.Settings.AttachmentHelperSettingsPanel();
            this.SuspendLayout();
            // 
            // btnSaveSetting
            // 
            this.btnSaveSetting.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSaveSetting.Location = new System.Drawing.Point(953, 420);
            this.btnSaveSetting.Name = "btnSaveSetting";
            this.btnSaveSetting.Size = new System.Drawing.Size(75, 23);
            this.btnSaveSetting.TabIndex = 7;
            this.btnSaveSetting.Text = "OK";
            this.btnSaveSetting.UseVisualStyleBackColor = true;
            this.btnSaveSetting.Click += new System.EventHandler(this.btnSaveSetting_Click);
            // 
            // ctrlMenu
            // 
            this.ctrlMenu.BackColor = System.Drawing.SystemColors.Window;
            this.ctrlMenu.Items = new string[] {
        "Attachments Helper"};
            this.ctrlMenu.Location = new System.Drawing.Point(3, 1);
            this.ctrlMenu.Name = "ctrlMenu";
            this.ctrlMenu.SelectedItem = 0;
            this.ctrlMenu.Size = new System.Drawing.Size(165, 442);
            this.ctrlMenu.TabIndex = 9;
            this.ctrlMenu.MenuItemSelected += new System.EventHandler<HelperUtils.MenuItemSelectedEventArgs>(this.ctrlMenu_MenuItemSelected);
            this.ctrlMenu.Load += new System.EventHandler(this.ctrlMenu_Load);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(872, 420);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 13;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // ctrlAttachments
            // 
            this.ctrlAttachments.BackColor = System.Drawing.SystemColors.Window;
            this.ctrlAttachments.Location = new System.Drawing.Point(174, 1);
            this.ctrlAttachments.Name = "ctrlAttachments";
            this.ctrlAttachments.Size = new System.Drawing.Size(854, 413);
            this.ctrlAttachments.TabIndex = 14;
            this.ctrlAttachments.Visible = false;
            // 
            // SettingsFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1031, 445);
            this.Controls.Add(this.ctrlAttachments);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.ctrlMenu);
            this.Controls.Add(this.btnSaveSetting);
            this.Name = "SettingsFrm";
            this.Text = "OutlookHelper Settings";
            this.Load += new System.EventHandler(this.SettingsFrm_Load);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button btnSaveSetting;
        private HelperUtils.SideMenu ctrlMenu;
        private AttachmentHelperSettingsPanel ctrlAttachmentsPanel;
        private System.Windows.Forms.Button btnCancel;
        private AttachmentHelperSettingsPanel ctrlAttachments;
    }
}