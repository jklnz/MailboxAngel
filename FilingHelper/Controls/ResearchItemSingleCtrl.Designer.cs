﻿namespace FilingHelper.Controls
{
    partial class ResearchItemSingleCtrl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ResearchItemSingleCtrl));
            this.pnlContainer = new System.Windows.Forms.Panel();
            this.txtComment = new System.Windows.Forms.TextBox();
            this.picMailIcon = new System.Windows.Forms.PictureBox();
            this.lblSubject = new System.Windows.Forms.Label();
            this.txtBody = new System.Windows.Forms.RichTextBox();
            this.ctlToolStrip = new System.Windows.Forms.ToolStrip();
            this.btnOpenItem = new System.Windows.Forms.ToolStripButton();
            this.btnPintoBoard = new System.Windows.Forms.ToolStripButton();
            this.btnNote = new System.Windows.Forms.ToolStripButton();
            this.btnDeleteItem = new System.Windows.Forms.ToolStripButton();
            this.ctlToolTip = new System.Windows.Forms.ToolTip(this.components);
            this.pnlContainer.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picMailIcon)).BeginInit();
            this.ctlToolStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlContainer
            // 
            this.pnlContainer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlContainer.Controls.Add(this.txtComment);
            this.pnlContainer.Controls.Add(this.picMailIcon);
            this.pnlContainer.Controls.Add(this.lblSubject);
            this.pnlContainer.Controls.Add(this.txtBody);
            this.pnlContainer.Controls.Add(this.ctlToolStrip);
            this.pnlContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlContainer.Location = new System.Drawing.Point(0, 0);
            this.pnlContainer.Name = "pnlContainer";
            this.pnlContainer.Size = new System.Drawing.Size(309, 98);
            this.pnlContainer.TabIndex = 0;
            this.pnlContainer.Enter += new System.EventHandler(this.pnlContainer_Enter);
            this.pnlContainer.Leave += new System.EventHandler(this.pnlContainer_Leave);
            this.pnlContainer.MouseDown += new System.Windows.Forms.MouseEventHandler(this.pnlContainer_MouseDown);
            this.pnlContainer.MouseEnter += new System.EventHandler(this.pnlContainer_MouseEnter);
            this.pnlContainer.MouseLeave += new System.EventHandler(this.pnlContainer_MouseLeave);
            this.pnlContainer.Resize += new System.EventHandler(this.pnlContainer_Resize);
            // 
            // txtComment
            // 
            this.txtComment.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.txtComment.Location = new System.Drawing.Point(-1, 21);
            this.txtComment.Multiline = true;
            this.txtComment.Name = "txtComment";
            this.txtComment.Size = new System.Drawing.Size(277, 78);
            this.txtComment.TabIndex = 5;
            this.txtComment.Visible = false;
            this.txtComment.TextChanged += new System.EventHandler(this.txtComment_TextChanged);
            this.txtComment.Enter += new System.EventHandler(this.txtComment_Enter);
            this.txtComment.Leave += new System.EventHandler(this.txtComment_Leave);
            this.txtComment.Validating += new System.ComponentModel.CancelEventHandler(this.txtComment_Validating);
            // 
            // picMailIcon
            // 
            this.picMailIcon.Image = global::FilingHelper.Properties.Resources.icon_msg_unread;
            this.picMailIcon.Location = new System.Drawing.Point(2, 3);
            this.picMailIcon.Name = "picMailIcon";
            this.picMailIcon.Size = new System.Drawing.Size(17, 19);
            this.picMailIcon.TabIndex = 2;
            this.picMailIcon.TabStop = false;
            this.picMailIcon.Click += new System.EventHandler(this.picMailIcon_Click);
            // 
            // lblSubject
            // 
            this.lblSubject.AutoSize = true;
            this.lblSubject.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSubject.Location = new System.Drawing.Point(17, 5);
            this.lblSubject.Name = "lblSubject";
            this.lblSubject.Size = new System.Drawing.Size(41, 13);
            this.lblSubject.TabIndex = 0;
            this.lblSubject.Text = "label1";
            this.lblSubject.MouseDown += new System.Windows.Forms.MouseEventHandler(this.lblSubject_MouseDown);
            // 
            // txtBody
            // 
            this.txtBody.BackColor = System.Drawing.SystemColors.Control;
            this.txtBody.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtBody.Location = new System.Drawing.Point(-1, 21);
            this.txtBody.Name = "txtBody";
            this.txtBody.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None;
            this.txtBody.Size = new System.Drawing.Size(309, 76);
            this.txtBody.TabIndex = 7;
            this.txtBody.Text = "";
            // 
            // ctlToolStrip
            // 
            this.ctlToolStrip.Dock = System.Windows.Forms.DockStyle.Right;
            this.ctlToolStrip.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.ctlToolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnOpenItem,
            this.btnPintoBoard,
            this.btnNote,
            this.btnDeleteItem});
            this.ctlToolStrip.Location = new System.Drawing.Point(283, 0);
            this.ctlToolStrip.Name = "ctlToolStrip";
            this.ctlToolStrip.Size = new System.Drawing.Size(24, 96);
            this.ctlToolStrip.TabIndex = 4;
            this.ctlToolStrip.Text = "toolStrip1";
            // 
            // btnOpenItem
            // 
            this.btnOpenItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnOpenItem.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenItem.Image")));
            this.btnOpenItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnOpenItem.Name = "btnOpenItem";
            this.btnOpenItem.Size = new System.Drawing.Size(21, 20);
            this.btnOpenItem.Text = "toolStripButton1";
            this.btnOpenItem.ToolTipText = "Open mail item";
            this.btnOpenItem.Click += new System.EventHandler(this.btnOpenItem_Click);
            // 
            // btnPintoBoard
            // 
            this.btnPintoBoard.CheckOnClick = true;
            this.btnPintoBoard.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnPintoBoard.Image = ((System.Drawing.Image)(resources.GetObject("btnPintoBoard.Image")));
            this.btnPintoBoard.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnPintoBoard.Name = "btnPintoBoard";
            this.btnPintoBoard.Size = new System.Drawing.Size(21, 20);
            this.btnPintoBoard.Text = "toolStripButton1";
            this.btnPintoBoard.ToolTipText = "Pin this mail item to the board";
            this.btnPintoBoard.Click += new System.EventHandler(this.btnPintoBoard_Click);
            // 
            // btnNote
            // 
            this.btnNote.CheckOnClick = true;
            this.btnNote.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnNote.Image = ((System.Drawing.Image)(resources.GetObject("btnNote.Image")));
            this.btnNote.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnNote.Name = "btnNote";
            this.btnNote.Size = new System.Drawing.Size(21, 20);
            this.btnNote.Text = "toolStripButton1";
            this.btnNote.ToolTipText = "Show note for this item";
            this.btnNote.Click += new System.EventHandler(this.btnNote_Click);
            // 
            // btnDeleteItem
            // 
            this.btnDeleteItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnDeleteItem.Image = ((System.Drawing.Image)(resources.GetObject("btnDeleteItem.Image")));
            this.btnDeleteItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnDeleteItem.Name = "btnDeleteItem";
            this.btnDeleteItem.Size = new System.Drawing.Size(21, 20);
            this.btnDeleteItem.Text = "toolStripButton2";
            this.btnDeleteItem.ToolTipText = "Remove this mail item from the list";
            this.btnDeleteItem.Click += new System.EventHandler(this.btnDeleteItem_Click);
            // 
            // ctlToolTip
            // 
            this.ctlToolTip.AutoPopDelay = 5000;
            this.ctlToolTip.InitialDelay = 0;
            this.ctlToolTip.IsBalloon = true;
            this.ctlToolTip.ReshowDelay = 100;
            // 
            // ResearchItemSingleCtrl
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.pnlContainer);
            this.Name = "ResearchItemSingleCtrl";
            this.Size = new System.Drawing.Size(309, 98);
            this.pnlContainer.ResumeLayout(false);
            this.pnlContainer.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picMailIcon)).EndInit();
            this.ctlToolStrip.ResumeLayout(false);
            this.ctlToolStrip.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlContainer;
        private System.Windows.Forms.Label lblSubject;
        private System.Windows.Forms.PictureBox picMailIcon;
        private System.Windows.Forms.ToolStrip ctlToolStrip;
        private System.Windows.Forms.ToolStripButton btnOpenItem;
        private System.Windows.Forms.ToolStripButton btnDeleteItem;
        private System.Windows.Forms.TextBox txtComment;
        private System.Windows.Forms.ToolStripButton btnPintoBoard;
        private System.Windows.Forms.ToolStripButton btnNote;
        private System.Windows.Forms.RichTextBox txtBody;
        private System.Windows.Forms.ToolTip ctlToolTip;
    }
}
