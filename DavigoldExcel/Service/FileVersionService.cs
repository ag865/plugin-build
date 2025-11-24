using DavigoldExcel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DavigoldExcel.Service
{
    public class FileVersionService
    {
        public static async Task<bool> UploadExcelFile(byte[] workbookData, string templateFileId, string fileName, string fileType)                             
        {
            string tenantUrl = AuthService.GetStoredTenantUrl();
            User user = AuthService.GetStoredUser();                                       
            string token = AuthService.GetStoredToken();
            int? TenantId = user.TenantId;     


            UriBuilder imageUrlBuilder = new UriBuilder(tenantUrl)
            {
                Path = "/api/image"
            };

            UriBuilder uriBuilder = new UriBuilder(tenantUrl)
            {
                Path = "/api/addin/sync-file"
            };

            ApiService _apiService = new ApiService
            {
                Token = token
            };

            var response = await _apiService.SyncFileExcel(uriBuilder.ToString(), new { tenantId = TenantId, templateFileId = templateFileId, fileName = fileName, fileType, imageUrl = imageUrlBuilder.ToString() + "/", type = "excel" }, workbookData);

            return response;
        }

        public static async Task<bool> UploadPresentationFile(byte[] workbookData, string templateFileId, string fileName, string fileType)
        {
            string tenantUrl = AuthService.GetStoredTenantUrl();
            User user = AuthService.GetStoredUser();
            string token = AuthService.GetStoredToken();
            int? TenantId = user.TenantId;


            UriBuilder imageUrlBuilder = new UriBuilder(tenantUrl)
            {
                Path = "/api/image"
            };

            UriBuilder uriBuilder = new UriBuilder(tenantUrl)
            {
                Path = "/api/addin/sync-file"
            };

            ApiService _apiService = new ApiService
            {
                Token = token
            };

            var response = await _apiService.SyncFilePresentation(uriBuilder.ToString(), new { tenantId = TenantId, templateFileId = templateFileId, fileName = fileName, fileType, imageUrl = imageUrlBuilder.ToString() + "/", type = "ppt" }, workbookData);

            return response;
        }

        public static async Task<bool> UploadWordFile(byte[] workbookData, string templateFileId, string fileName, string fileType)
        {
            string tenantUrl = AuthService.GetStoredTenantUrl();
            User user = AuthService.GetStoredUser();
            string token = AuthService.GetStoredToken();
            int? TenantId = user.TenantId;


            UriBuilder imageUrlBuilder = new UriBuilder(tenantUrl)
            {
                Path = "/api/image"
            };

            UriBuilder uriBuilder = new UriBuilder(tenantUrl)
            {
                Path = "/api/addin/sync-file"
            };

            ApiService _apiService = new ApiService
            {
                Token = token
            };

            var response = await _apiService.SyncFileWord(uriBuilder.ToString(), new { tenantId = TenantId, templateFileId = templateFileId, fileName = fileName, fileType, imageUrl = imageUrlBuilder.ToString() + "/", type = "word" }, workbookData);

            return response;
        }
    }
}
 