using DavigoldExcel.Models;
using DavigoldExcel.Service;
using System;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Security.Cryptography;
namespace DavigoldExcel.Controls
{
    /// <summary>
    /// Interaction logic for UserTokenForm.xaml
    /// </summary>
    public partial class UserTokenForm : UserControl
    {
        ApiService apiService;
        public UserTokenForm()
        {
            InitializeComponent();
            apiService = new ApiService();
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            //  Login Button Click implementation
            string currentToken = TokenTextBox.Text.ToString(); 
            string currentTenantUrl = TenantUrlTextBox.Text.ToString();

            if (!String.IsNullOrWhiteSpace(currentToken) && !String.IsNullOrWhiteSpace(currentTenantUrl) )
            {
                LoginButton.IsEnabled = false;
                LoginButton.Content = "Loading...";
                UriBuilder uriBuilder = new UriBuilder(currentTenantUrl)
                {
                    Path = "/api/auth/token"
                };

                try
                {
                    var response = await apiService.PostJsonData<TokenResponse>(uriBuilder.ToString(), new { token = currentToken });
                    LoginButton.Content = "Login";
                    LoginButton.IsEnabled = true;

                    if (response != null)
                    {
                        AuthService.StoreToken(currentToken);
                        AuthService.StoreUser(response.user);
                        AuthService.StoreTenantUrl(currentTenantUrl);

                        Globals.ThisAddIn.InitializeAuth();
                        if (!UTILS.showSidePanel)
                        {
                            Globals.ThisAddIn.ToggleTaskPane();
                        }
                        MessageBox.Show("You are successfully logged in");
                    }
                }
                catch (System.Net.Http.HttpRequestException ex)
                {
                    LoginButton.Content = "Login";
                    LoginButton.IsEnabled = true;
                    
                    // Show the error message from the server, or a default message if none available
                    string errorMessage = !string.IsNullOrWhiteSpace(ex.Message) 
                        ? ex.Message 
                        : "Invalid Credentials!";
                    MessageBox.Show(errorMessage, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                catch (Exception ex)
                {
                    LoginButton.Content = "Login";
                    LoginButton.IsEnabled = true;
                    
                    // Show the error message from the exception
                    string errorMessage = !string.IsNullOrWhiteSpace(ex.Message) 
                        ? ex.Message 
                        : "An error occurred while attempting to login.";
                    MessageBox.Show(errorMessage, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}
