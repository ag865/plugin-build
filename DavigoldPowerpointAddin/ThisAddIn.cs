using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using DavigoldExcel;
using DavigoldExcel.Service;
using DavigoldExcel.Models;
using System.IO;
using System.Windows;

namespace DavigoldPowerpointAddin
{
    public partial class ThisAddIn
    {
        PPMainWindow mainWindow;
        private bool isLoggedIn;
        private User user;
        private string token;
        private Dictionary<string, CustomTaskPane> customPanes = new Dictionary<string, CustomTaskPane>();

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

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.WindowActivate += WindowActivate;
            Application.PresentationSave += Application_PresentationSave;
        }

        private async void Application_PresentationSave(Presentation Pres)
        {
            //string currentWorkbookName = Pres.Name;
            //List<string> nameSegments = currentWorkbookName.Split('-').ToList();
            //if (nameSegments.Count >= 4 && nameSegments[0] == "id")
            //{
            //    string currentTemplateFileId = nameSegments[1];
            //    string fileType = nameSegments[2];
            //    string currentFileName = nameSegments[nameSegments.Count - 1];

            //    try
            //    {
            //        byte[] presentationData = File.ReadAllBytes(Pres.FullName);
            //        // Send the workbook data to the server
            //        var result = await FileVersionService.UploadPresentationFile(presentationData, currentTemplateFileId, currentFileName, fileType);
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show(ex.Message);
            //    }
            //}

            //// Convert the boolean value to MsoTriState
            //Office.MsoTriState savedState = Office.MsoTriState.msoTrue;

            //// Set the saved state of the presentation
            //Pres.Saved = savedState;
        }

        private void WindowActivate(Presentation Pres, DocumentWindow Wn)
        {
            string caption = Application.ActiveWindow.Caption;
            // Check if a CustomPane exists for the active window
            if (!customPanes.ContainsKey(caption))
            {

                mainWindow = new PPMainWindow();
                CustomTaskPane mainTaskPane = this.CustomTaskPanes.Add(mainWindow, "ONE");

                RefetchUser();

                // Store the CustomPane associated with the window
                customPanes.Add(caption, mainTaskPane);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Application.WindowActivate -= WindowActivate;
            Application.PresentationSave -= Application_PresentationSave;
        }

        private void PresentationOpen(Presentation Pres)
        {
            string caption = Application.ActiveWindow.Caption;
            // Check if a CustomPane exists for the active window
            if (!customPanes.ContainsKey(caption))
            {

                mainWindow = new PPMainWindow();
                CustomTaskPane mainTaskPane = this.CustomTaskPanes.Add(mainWindow, "Davigold Panel");

                RefetchUser();

                // Store the CustomPane associated with the window
                customPanes.Add(caption, mainTaskPane);
            }
        }

        private CustomTaskPane GetCurrentTaskPane()
        {
            string caption = (string)Application.ActiveWindow.Caption;
            CustomTaskPane ctp = customPanes[caption];
            if (ctp != null)
            {
                return ctp;
            }
            return null;
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
                mainTaskPane.Width = 650;
            }

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
                Globals.Ribbons.DavigoldPPRibon.ShowHideButton.Visible = currentUser.ShowHideShowFields;
                Globals.Ribbons.DavigoldPPRibon.LoginLogoutButton.Label = "Logout";
            }
            else
            {
                isLoggedIn = false;
                Globals.Ribbons.DavigoldPPRibon.ShowHideButton.Visible = false;
                Globals.Ribbons.DavigoldPPRibon.LoginLogoutButton.Label = "Login";
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

        public string GetToken()
        {
            return this.token;
        }

        public User GetUser()
        {
            return this.user;
        }

        public bool IsLoggedIn()
        {
            return this.isLoggedIn;
        }
    }
}
