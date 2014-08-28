using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace save_to_google_drive
{
    public partial class SaveToGoogleDriveAddIn
    {
        internal static Outlook.Explorer CurrentExplorer { get; set; }

        private void SaveToGoogleDriveAddIn_Startup(object sender, System.EventArgs e)
        {
            CurrentExplorer = this.Application.ActiveExplorer();

            CurrentExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(CurrentExplorer_Event);
        }


        private void CurrentExplorer_Event()
        {
            
            Outlook.MAPIFolder selectedFolder =
                this.Application.ActiveExplorer().CurrentFolder;
            String expMessage = "Your current folder is "
                + selectedFolder.Name + ".\n";
            String itemMessage = "Item is unknown.";
            try
            {
                if (this.Application.ActiveExplorer().Selection.Count > 0)
                {
                    Object selObject = this.Application.ActiveExplorer().Selection[1];
                    if (selObject is Outlook.MailItem)
                    {
                        Outlook.MailItem mailItem =
                            (selObject as Outlook.MailItem);
                        itemMessage = "The item is an e-mail message." +
                            " The subject is " + mailItem.Subject + ".";
                        //mailItem.Display(false);
                    }
                }
            } catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        private void SaveToGoogleDriveAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(SaveToGoogleDriveAddIn_Startup);
            this.Shutdown += new System.EventHandler(SaveToGoogleDriveAddIn_Shutdown);
        }
        
        #endregion
    }
}
