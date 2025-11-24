using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DavigoldExcel.Service;
using DavigoldExcel.Models;
using System.Globalization;
using System.Windows.Controls;
using System.IO;
using Microsoft.Office.Interop.Word;
using DavigoldExcel;

namespace DavigoldWordAddin
{
    public partial class WordRibbon
    {
        private void WordRibbon_Load(object sender, RibbonUIEventArgs e)
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

        private void SyncButton_Click(object sender, RibbonControlEventArgs e)
        {
            Document Doc = Globals.ThisAddIn.Application.ActiveDocument;

            string currentDocumentName = Doc.Name;
            var oldPath = Doc.FullName;
            List<string> nameSegments = currentDocumentName.Split('-').ToList();

            if (nameSegments.Count >= 4 && nameSegments[0] == "id")
            {
                string currentTemplateFileId = nameSegments[1];
                string fileType = nameSegments[2];
                string currentFileName = nameSegments[nameSegments.Count - 1];

                string tempFilePath = Path.Combine(Doc.Path, Guid.NewGuid() + ".docx");
                string secondTempFilePath = Path.Combine(Doc.Path, Guid.NewGuid() + ".docx");

                try
                {
                    // Save the document to a temporary location
                    Doc.SaveAs2(tempFilePath, WdSaveFormat.wdFormatXMLDocument);
                    Doc.SaveAs2(secondTempFilePath, WdSaveFormat.wdFormatXMLDocument);

                    // Read the temporary file
                    byte[] documentData = File.ReadAllBytes(tempFilePath);
                    // byte[] documentData = File.ReadAllBytes(Doc.FullName);

                    // Send the document data to the server
                    var result = FileVersionService.UploadWordFile(documentData, currentTemplateFileId, currentFileName, fileType).GetAwaiter().GetResult();


                    if(File.Exists(oldPath))
                    {
                        File.Delete(oldPath);
                    }

                    if (File.Exists(tempFilePath))
                    {
                        File.Delete(tempFilePath);
                    }

                    Doc.SaveAs2(oldPath, WdSaveFormat.wdFormatXMLDocument);

                    if (File.Exists(secondTempFilePath)) 
                    { 
                         File.Delete(secondTempFilePath);
                    }

                    MessageBox.Show("Version Published Successfully!");

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
