using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using DavigoldExcel.Service;
using System.Windows;
using DavigoldExcel.Models;
using Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Microsoft.Office.Core;
using System.Windows.Forms.Design;
using Microsoft.Office.Tools;
using CustomTaskPane = Microsoft.Office.Tools.CustomTaskPane;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.IO;

namespace DavigoldExcel
{
  public partial class ThisAddIn
  {
    MainWindow mainWindow;
    private bool isLoggedIn;
    private User user;
    private string token;

    private Dictionary<string, CustomTaskPane> customPanes = new Dictionary<string, CustomTaskPane>();

    private void ThisAddIn_Startup(object sender, System.EventArgs e)
    {
      Application.WorkbookActivate += Application_WorkbookOpen;
      Application.WorkbookAfterSave += Application_WorkbookAfterSave;
    }

    private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
    {
      Application.WorkbookActivate -= Application_WorkbookOpen;
      Application.WorkbookAfterSave -= Application_WorkbookAfterSave;
    }

    private async void Application_WorkbookAfterSave(Workbook Wb, bool Success)
    {
      //string currentWorkbookName = Wb.Name;
      //List<string> nameSegments = currentWorkbookName.Split('-').ToList();
      //if(nameSegments.Count >= 4 && nameSegments[0] == "id" ) {
      //    string currentTemplateFileId = nameSegments[1];
      //    string fileType = nameSegments[2];
      //    string currentFileName = nameSegments[nameSegments.Count - 1];

      //    string tempFilePath = Path.GetTempFileName();
      //    Wb.SaveCopyAs(tempFilePath);

      //    try
      //    {
      //        byte[] workbookData;
      //        // Read the temporary file into a memory stream
      //        using (FileStream fileStream = new FileStream(tempFilePath, FileMode.Open, FileAccess.Read))
      //        {
      //            MemoryStream memoryStream = new MemoryStream();
      //            fileStream.CopyTo(memoryStream);

      //            workbookData = memoryStream.ToArray();
      //        }
      //        // Send the workbook data to the server
      //        var result = await FileVersionService.UploadExcelFile(workbookData, currentTemplateFileId, currentFileName, fileType);
      //    }
      //    catch (Exception ex)
      //    {
      //        MessageBox.Show(ex.Message);
      //    }
      //}
    }

    void Application_WorkbookOpen(Workbook workbook)
    {
      string caption = (string)Application.ActiveWindow.Caption;
      // Check if a CustomPane exists for the active window
      if (!customPanes.ContainsKey(caption))
      {

        mainWindow = new MainWindow();
        CustomTaskPane mainTaskPane = this.CustomTaskPanes.Add(mainWindow, "ONE");

        RefetchUser();

        // Store the CustomPane associated with the window
        customPanes.Add(caption, mainTaskPane);
      }
    }

    public CustomTaskPane GetCurrentTaskPane()
    {
      string caption = (string)Application.ActiveWindow.Caption;
      return customPanes[caption];
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


    public void InitializeAuth()
    {
      string token = AuthService.GetStoredToken();
      User currentUser = AuthService.GetStoredUser();

      if (currentUser != null && token != null)
      {
        isLoggedIn = true;
        this.user = currentUser;
        this.token = token;

        Globals.Ribbons.DavigoldRibbon.ImportButton.Visible = currentUser.ShowDownloadButton || UTILS.showImportButton;
        Globals.Ribbons.DavigoldRibbon.UploadButton.Visible = currentUser.ShowUploadButton || UTILS.showUploadButton;
        Globals.Ribbons.DavigoldRibbon.ShowHideButton.Visible = currentUser.ShowHideShowFields || UTILS.showShowHideButton;

        Globals.Ribbons.DavigoldRibbon.LogoutButton.Label = "Logout";
      }
      else
      {
        isLoggedIn = false;
        Globals.Ribbons.DavigoldRibbon.ImportButton.Visible = false;
        Globals.Ribbons.DavigoldRibbon.UploadButton.Visible = false;
        Globals.Ribbons.DavigoldRibbon.ShowHideButton.Visible = false;
        Globals.Ribbons.DavigoldRibbon.LogoutButton.Label = "Login";
      }

      UpdateMainWindowPage(isLoggedIn);
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

      // AuthService.Logout();
      InitializeAuth();
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
