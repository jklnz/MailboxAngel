using System;
using System.Linq;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Threading;
using Microsoft.Office.Tools;
using HelperUtils;
using AttachmentManager;
using FilingHelper.Controls;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Office.Interop.Outlook;
using MailboxAngel.OutlookCommon;

namespace FilingHelper
{
    /// <summary>
    /// Mailbox Angel
    /// ----------------------------------------------------------
    /// (c) Shai Shulman, 2021
    /// MIT License subject to "Commons Clause” License Condition v1.0
    /// ----------------------------------------------------------
    /// The addin provides additional functionalities for handling multiple mail folders, moving mail items into folders and handling attachments
    /// - Find folder by typing the first letter of its name
    /// - Move between folders with same no on different PSTs (next/previous)
    /// - Display list of recenty used folders
    /// - Show suggestion one where to file a mail items based on recent folders, other items in the conversation, and otehr items with the same sender
    /// - Attachment helper, allowing attachments manipulations directly from the composing windows:
    ///     - Change file name of attachments
    ///     - Change order of attachments
    ///     - Accept all changes in an attachment (for Word documents)
    ///     - Create zip file with the attachments and attach to the mail item
    /// </summary>
    public partial class ThisAddIn
    {
        const int HISTORY_SIZE = 15;
        const string APPLICATION_CAPTION = "Mailbox Angel";
        const string FOLDER_CHANGED_MESSAGE = "Folder Changed to";
        const string ITEM_MOVED_MESSAGE = "Mail Item(s) Moved to";
        Outlook.Inspectors inspectors;
        Outlook.Explorers explorers;
        private Dictionary<Outlook.Explorer, ExplorerWrapper> folderPanelsWrapper = new Dictionary<Outlook.Explorer, ExplorerWrapper>();
        //private OutlookWindowStore<FolderArchiver> folderArchivers = new OutlookWindowStore<FolderArchiver>();
        private OutlookWindowStore<AttachmentManager.AttachmentManager> attachmentManagers = new OutlookWindowStore<AttachmentManager.AttachmentManager>();
        //private UserControlStore<FolderArchiverCtrl> folderArchiverCtrls = new UserControlStore<FolderArchiverCtrl>();
        private UserControlStore<Controls.AttachmentsPaneCtrl> attachmentPaneCtrls = new UserControlStore<Controls.AttachmentsPaneCtrl>();
        //private UserControlStore<ResearchPanelCtrl> researchPaneCtril = new UserControlStore<ResearchPanelCtrl>();

           
        
        public Dictionary<Outlook.Explorer, ExplorerWrapper> FolderPanelsWrapper
        {
            get
            {
                return folderPanelsWrapper;
            }
        }

        private bool _updateAttachmentsOnAddRemove = true;
        public bool UpdateAttachmentsOnAddRemove
        {
            get { return _updateAttachmentsOnAddRemove; }
            set { _updateAttachmentsOnAddRemove = value; }
        }

      

        /// <summary>
        /// Initialize attachment manager for a mail item
        /// </summary>
        /// <param name="inspector">Inspector containing the mail item</param>
        public void AttachmentManager(Outlook.Inspector inspector)
        {
            if (!(inspector.CurrentItem is Outlook.MailItem))
                return;
            if (attachmentPaneCtrls[inspector]!=null && attachmentPaneCtrls[inspector].Visible)
            {
                HideAttachmentManager(inspector,false);
                return;
            }
            AttachmentManager.AttachmentManager manager = new AttachmentManager.AttachmentManager(inspector.CurrentItem);
            manager.AttachmentsFinished += Manager_AttachmentsFinished;
            attachmentManagers[inspector] = manager;
            if (attachmentPaneCtrls[inspector] == null)
                attachmentPaneCtrls.Insert(new AttachmentsPaneCtrl(manager), "Attachment Manager", inspector);
            CustomTaskPane pane = attachmentPaneCtrls[inspector];
            ((AttachmentsPaneCtrl)pane.Control).TotalWidth = pane.Width;
            pane.Visible = true;
            pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop;
            manager.UIElement = pane;
            List<AttachmentCommand> attachments = (manager.getAttachments());

            Outlook.MailItem message = (Outlook.MailItem)inspector.CurrentItem;
            message.AttachmentAdd += ThisAddIn_AttachmentAdd;
            message.AttachmentRemove += Message_AttachmentRemove;
            ((AttachmentsPaneCtrl)pane.Control).Fill(attachments);
            ((AttachmentsPaneCtrl)pane.Control).AttachmentsUpdated += AttachmentPaneCtrl_AttachmentsUpdated;
            ((AttachmentsPaneCtrl)pane.Control).Resize += ((s, e) =>
            {
                //((AttachmentsPaneCtrl)pane.Control).TotalWidth = pane.Width;
                //if (pane.Height > ((AttachmentsPaneCtrl)pane.Control).MaxHeight)
                //    pane.Height = ((AttachmentsPaneCtrl)pane.Control).MaxHeight;
                //else
                //    ((AttachmentsPaneCtrl)pane.Control).TotalHeight = pane.Height;
            });
            pane.Height = ((AttachmentsPaneCtrl)pane.Control).TotalHeight;
        }

        private void AttachmentPaneCtrl_AttachmentsUpdated(object sender, AttachmentsUpdatedEventArgs e)
        {
            CustomTaskPane pane = (e.UIElement) as CustomTaskPane;
            pane.Height = ((AttachmentsPaneCtrl)pane.Control).TotalHeight;
        }

        /// <summary>
        /// Hide the attachment manager for a specific mail item
        /// </summary>
        /// <param name="inspector">Inspector containing the mail items</param>
        /// <param name="fRemove">Terminate pane if True, otherwise only hide pane</param>
        private void HideAttachmentManager(Outlook.Inspector inspector, bool fRemove=false)
        {


            if (!(inspector.CurrentItem is Outlook.MailItem))
                return;
            if (fRemove)
            {
                attachmentPaneCtrls.Remove(inspector);
                attachmentManagers.Remove(inspector);
            } else
            {
                if (attachmentPaneCtrls[inspector] != null)
                    attachmentPaneCtrls[inspector].Visible=false;
            }
        }

        /// <summary>
        /// Listener for adding an attachment to a mail item (will add to relevant attachment manager if open)
        /// </summary>
        /// <param name="Attachment">Attachment object added</param>
        private void ThisAddIn_AttachmentAdd(Outlook.Attachment Attachment)
        {
            if (_updateAttachmentsOnAddRemove)
            {
                Outlook.Inspector inspector = ((Outlook.MailItem)Attachment.Parent).GetInspector;
                AttachmentManager.AttachmentManager manager = attachmentManagers[inspector];
                ((AttachmentsPaneCtrl)((CustomTaskPane)manager.UIElement).Control).Add(new ExistingAttachmentCommand(Attachment));
            }
        }
        /// <summary>
        /// Listener for removing an attachment to a mail item (will remove from relevant attachment manager if open)
        /// </summary>
        /// <param name="Attachment">Attachment object added</param>
        private void Message_AttachmentRemove(Outlook.Attachment Attachment)
        {
            if (_updateAttachmentsOnAddRemove)
            {
                Outlook.Inspector inspector = ((Outlook.MailItem)Attachment.Parent).GetInspector;
                AttachmentManager.AttachmentManager manager = attachmentManagers[inspector];
                ((AttachmentsPaneCtrl)((CustomTaskPane)manager.UIElement).Control).Remove(Attachment);
            }
        }

        private void Manager_AttachmentsFinished(object sender, AttachmentsFinishedEventArgs e)
        {
            CustomTaskPane pane= (CustomTaskPane)e.UIElement;
            pane.Visible = false;
            pane.Height = ((AttachmentsPaneCtrl)pane.Control).TotalHeight;
        }


       

        public DialogResult CustomMessageBox(string message,MessageBoxButtons buttons,MessageBoxIcon icon)
        {
            return MessageBox.Show(message, APPLICATION_CAPTION, buttons, icon);
        }

        /// <summary>
        /// Load data and initialize all objects on addin startup
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector += Inspectors_NewInspector;
            explorers = this.Application.Explorers;
            explorers.NewExplorer += Explorers_NewExplorer;
            foreach (Outlook.Explorer explorer in explorers)
            {
                Explorers_NewExplorer(explorer);
            }

            this.Application.ItemSend += Application_ItemSend;

            ((Outlook.ApplicationEvents_11_Event)Application).Quit += ThisAddIn_Quit;
        }


        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            MailItem mail = (MailItem)Item;
            if (mail.MessageClass == "IPM.Note")
                (new SignaturesService()).ApplyCustomSignature(mail);
        }

        private void Items_ItemAdd(object Item)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Save settings when addin (and Outlook) is terminated
        /// </summary>
        private void ThisAddIn_Quit()
        {

        }

        /// <summary>
        /// Add necessary objects to a newly created explorer
        /// </summary>
        /// <param name="Explorer">Newly created explorer</param>
        private void Explorers_NewExplorer(Outlook.Explorer Explorer)
        {
            
            Explorer.InlineResponse += Explorer_InlineResponse;
            Explorer.InlineResponseClose += Explorer_InlineResponseClose;
        }

        private void Explorer_InlineResponseClose()
        {
            //throw new NotImplementedException();
        }

        private void Explorer_InlineResponse(object Item)
        {
           // throw new NotImplementedException();
            
        }



        private void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            if (Properties.AddinSettings.Default.MailHistoryAddMode== MailHistoryAddMode.InspectorOpened
                && Inspector.CurrentItem is Outlook.MailItem)
            {
                Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            explorers.NewExplorer -= Explorers_NewExplorer;
            explorers = null;
            folderPanelsWrapper = null;
            //archivePanelsWrapper = null;
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
