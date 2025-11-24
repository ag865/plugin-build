using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DavigoldExcel.Service;
using DavigoldExcel.Models;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Windows.Forms;
using MessageBox = System.Windows.MessageBox;
using System.Globalization;
using DavigoldExcel;
using System.IO;

namespace DavigoldPowerpointAddin
{
    public partial class DavigoldPPRibon
    {
        private void DavigoldPPRibon_Load(object sender, RibbonUIEventArgs e)
        {
            var version = typeof(ThisAddIn).Assembly.GetName().Version;
            DateTime _dateTime = new DateTime(2025, 10, 13);
            VersionLabel.Label = $"Version: {version}";
            UpdatedOnLabel.Label = $"Last Updated on {_dateTime.ToString("dd MMMM yyyy", CultureInfo.InvariantCulture)}";
            if (!UTILS.showSidePanel)
            {
                ShowHideButton.Visible = false;
            }
        }

        private void ShowHideButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ToggleTaskPane();
        }

        private void LoginLogoutButton_Click(object sender, RibbonControlEventArgs e)
        {
            User CurrentUser = AuthService.GetStoredUser();

            if (CurrentUser != null)
            {
                AuthService.Logout();
                Globals.ThisAddIn.InitializeAuth();
                MessageBox.Show("You are logged out from ONE plugin");
            }
            else
            {
                Globals.ThisAddIn.ToggleTaskPane();
            }
        }

        private void updateLinkButton_Click(object sender, RibbonControlEventArgs e)
        {

        }


        private async void SyncButton_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            pptApp.ActivePresentation.Save();
            Presentation Pres = pptApp.ActivePresentation;
            string currentWorkbookName = Pres.Name;
            List<string> nameSegments = currentWorkbookName.Split('-').ToList();
            if (nameSegments.Count >= 4 && nameSegments[0] == "id")
            {
                string currentTemplateFileId = nameSegments[1];
                string fileType = nameSegments[2];
                string currentFileName = nameSegments[nameSegments.Count - 1];

                try
                {
                    byte[] presentationData = File.ReadAllBytes(Pres.FullName);
                    // Send the workbook data to the server
                    var result = await FileVersionService.UploadPresentationFile(presentationData, currentTemplateFileId, currentFileName, fileType);
                    MessageBox.Show("Version Published Successfully!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            // Convert the boolean value to MsoTriState
            //Office.MsoTriState savedState = Office.MsoTriState.msoTrue;

            // Set the saved state of the presentation
            //Pres.Saved = savedState;
        }
    }
}
