using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;
using FilingHelper.Controls;
using HelperUtils;
using MailboxAngel.OutlookCommon;

namespace FilingHelper
{
    public partial class InspectorCustomRibbon
    {


        private void btnAttachments_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.AttachmentManager(Globals.ThisAddIn.Application.ActiveInspector());
        }

     

        private void btnItemReplyAll_Click(object sender, RibbonControlEventArgs e)
        {
            Inspector inspector = Globals.ThisAddIn.Application.ActiveInspector();
            if (inspector.CurrentItem is MailItem)
            {
                MailItem original = inspector.CurrentItem as MailItem;
                (new ResponseServices()).ReplyAttachments(original, true);
            }
        }

        private void btnItemReply_Click(object sender, RibbonControlEventArgs e)
        {
            Inspector inspector = Globals.ThisAddIn.Application.ActiveInspector();
            if (inspector.CurrentItem is MailItem)
            {
                MailItem original = inspector.CurrentItem as MailItem;
                (new ResponseServices()).ReplyAttachments(original, false);
            }

        }

    }
}
