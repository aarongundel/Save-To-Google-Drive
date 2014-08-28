using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;

namespace save_to_google_drive
{
    public partial class GoogleDrive
    {
        private void GoogleDrive_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void buttonGD_Click(object sender, RibbonControlEventArgs e)
        {
            Explorer exp = SaveToGoogleDriveAddIn.CurrentExplorer;
            try
            {
                if (exp.Selection.Count > 0)
                {
                    Object selObject = exp.Selection[1];
                    if (selObject is MailItem)
                    {
                        MailItem mailItem =
                            (selObject as MailItem);

                        // Okay, we're good to transfer.  Let's gooooooo.
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
    }
}
