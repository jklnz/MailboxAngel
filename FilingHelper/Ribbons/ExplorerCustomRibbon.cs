using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using HelperUtils;
using Microsoft.Office.Interop.Outlook;
using FilingHelper.Controls;
using MailboxAngel.OutlookCommon;
using System.Windows.Forms;

namespace FilingHelper
{
    public partial class ExplorerCustomRibbon
    {
        Controls.Settings.SettingsFrm _settingsForm;
  


        private void btnCloseAll_Click(object sender, RibbonControlEventArgs e)
        {
            (new HelperUtils.WindowManager()).CloseAll(Globals.ThisAddIn.Application);
        }

             


        private void btnItemReplyAll_Click(object sender, RibbonControlEventArgs e)
        {
            Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            if (explorer.Selection.Count > 0 && explorer.Selection[1] is MailItem)
            {
                MailItem original = explorer.Selection[1] as MailItem;
                (new ResponseServices()).ReplyAttachments(original, true);
            }

        }

        private void btnItemReply_Click(object sender, RibbonControlEventArgs e)
        {
            Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            if (explorer.Selection.Count > 0 && explorer.Selection[1] is MailItem)
            {
                MailItem original = explorer.Selection[1] as MailItem;
                (new ResponseServices()).ReplyAttachments(original, false);
            }

        }
       
    }
}

