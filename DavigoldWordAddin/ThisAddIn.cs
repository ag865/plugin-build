using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;
using DavigoldExcel.Models;
using Microsoft.Office.Tools;
using DavigoldExcel.Service;
using System.Windows.Forms;
using System.IO;

namespace DavigoldWordAddin
{
    public partial class ThisAddIn
    {
        WordMainWindow mainWindow;
        private bool isLoggedIn;
        private User user;
        private string token;
        private Dictionary<string, CustomTaskPane> customPanes = new Dictionary<string, CustomTaskPane>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.WindowActivate += WindowActive;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Application.WindowActivate -= WindowActive;
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

        private void WindowActive(Word.Document Doc, Window Wn)
        {
            string caption = Application.ActiveWindow.Caption;
            // Check if a CustomPane exists for the active window
            if (!customPanes.ContainsKey(caption))
            {

                mainWindow = new WordMainWindow();
                CustomTaskPane mainTaskPane = this.CustomTaskPanes.Add(mainWindow, "ONE");

                RefetchUser();

                // Store the CustomPane associated with the window
                customPanes.Add(caption, mainTaskPane);
            }
        }

        private void Application_DocumentBeforeSave(Word.Document Doc, ref bool saveASUI, ref bool Cancel)
        {
            Cancel = false;
            Doc.Save();

            string currentDocumentName = Doc.Name;
            List<string> nameSegments = currentDocumentName.Split('-').ToList();

            if (nameSegments.Count >= 4 && nameSegments[0] == "id")
            {
                string currentTemplateFileId = nameSegments[1];
                string fileType = nameSegments[2];
                string currentFileName = nameSegments[nameSegments.Count - 1];

                try
                {
                    byte[] documentData = File.ReadAllBytes(Doc.FullName);

                    // Send the document data to the server
                    var result = FileVersionService.UploadPresentationFile(documentData, currentTemplateFileId, currentFileName, fileType).GetAwaiter().GetResult();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // Mark the document as saved
            Doc.Saved = true;
        }


        public async void RefetchUser()
        {
            ApiService apiService = new ApiService();
            //  Login Button Click implementation
            string currentToken = AuthService.GetStoredToken();
            string currentTenantUrl = AuthService.GetStoredTenantUrl();

            if (!String.IsNullOrWhiteSpace(currentToken) && !String.IsNullOrWhiteSpace(currentTenantUrl))
            {
                UriBuilder uriBuilder = new UriBuilder(currentTenantUrl)
                {
                    Path = "/api/auth/token"
                };

                var response = await apiService.PostJsonData<TokenResponse>(uriBuilder.ToString(), new { token = currentToken });

                if (response != null)
                {

                    AuthService.StoreUser(response.user);
                    InitializeAuth();
                    return;
                }
            }

            AuthService.Logout();
            InitializeAuth();
        }

        public void InitializeAuth()
        {
            string token = AuthService.GetStoredToken();
            User currentUser = AuthService.GetStoredUser();

            if (currentUser != null && token != null)
            {
                isLoggedIn = true;
                this.user = currentUser;
                this.token = token;
                Globals.Ribbons.WordRibbon.ShowHideButton.Visible = currentUser.ShowHideShowFields;
                Globals.Ribbons.WordRibbon.LoginLogoutButton.Label = "Logout";
            }
            else
            {
                isLoggedIn = false;
                Globals.Ribbons.WordRibbon.ShowHideButton.Visible = false;
                Globals.Ribbons.WordRibbon.LoginLogoutButton.Label = "Login";
            }

            UpdateMainWindowPage(isLoggedIn);
        }

        public void UpdateMainWindowPage(bool isLogged)
        {
            if (isLogged)
            {
                mainWindow.ShowHomePage();
            }
            else
            {
                mainWindow.ShowTokenPage();
            }
        }

        public void ToggleTaskPane()
        {
            CustomTaskPane mainTaskPane = GetCurrentTaskPane();
            if (mainTaskPane.Visible)
            {
                mainTaskPane.Visible = false;
            }
            else
            {
                mainTaskPane.Visible = true;
                mainTaskPane.Width = 475;
            }

        }

        private CustomTaskPane GetCurrentTaskPane()
        {
            string caption = (string)Application.ActiveWindow.Caption;
            return customPanes[caption];
        }

        public User GetUser()
        {
            return this.user;
        }
    }
}
