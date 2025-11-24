using DavigoldExcel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;


namespace DavigoldExcel.Service
{
    public class ImportDataService
    {
        public async Task<List<T>> ImportData<T>(List<string> Columns, string Fund, string Module, string SubModule, string language = "EN", string date = null)
        {
            string tenantUrl = AuthService.GetStoredTenantUrl();
            User user = AuthService.GetStoredUser();
            string token = AuthService.GetStoredToken();
            int? TenantId = user.TenantId;

            UriBuilder uriBuilder = new UriBuilder(tenantUrl)
            {
                Path = "/api/addin/import-data"
            };

            ApiService _apiService = new ApiService
            {
                Token = token
            };

            var requestData = new
            {
                module = Module,
                subModule = SubModule,
                fund = Fund,
                columns = Columns,
                tenantId = TenantId,
                language,
                date
            };

            try
            {
                var response = await _apiService.PostJsonData<List<T>>(uriBuilder.ToString(), requestData);
                return response;
            }

            catch(Exception e)
            {
                MessageBox.Show(e.Message.ToString(), "Error", MessageBoxButton.OK);
                return new List<T>();
            }


        }
    }
}
