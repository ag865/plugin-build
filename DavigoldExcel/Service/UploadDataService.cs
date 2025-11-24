using DavigoldExcel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DavigoldExcel.Service
{
    public class UploadDataService
    {
        public async Task UploadData<T>(List<dynamic> data, string fundName, string date, string module, string mainModule, string language = "EN")
        {
            string tenantUrl = AuthService.GetStoredTenantUrl();
            User user = AuthService.GetStoredUser();
            string token = AuthService.GetStoredToken();
            int? TenantId = user.TenantId;

            UriBuilder uriBuilder = new UriBuilder(tenantUrl)
            {
                Path = "/api/addin/upload-data"
            };

            ApiService _apiService = new ApiService
            {
                Token = token
            };

            await _apiService.PostJsonData<T>(uriBuilder.ToString(), new { module, mainModule, data, fundName, date, tenantId = TenantId, language });

        }
    }
}
