using DavigoldExcel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DavigoldExcel.Service
{
    public class AuthService
    {

        private static readonly string TokenKey = "token";
        private static readonly string UserKey = "user";
        private static readonly string TenantUrlKey = "tenantUrl";

        public static void StoreToken(string token)
        {
            StorageService.StoreData(TokenKey, token);
        }

        public static string GetStoredToken()
        {
            return StorageService.GetStoredData<string>(TokenKey);
        }

        public static void StoreUser(User user)
        {
            StorageService.StoreData(UserKey, user);
        }

        public static User GetStoredUser()
        {
            return StorageService.GetStoredData<User>(UserKey);
        }

        public static void StoreTenantUrl(string tenantUrl)
        {
            StorageService.StoreData(TenantUrlKey, tenantUrl);
        }

        public static string GetStoredTenantUrl()
        {
            return StorageService.GetStoredData<string>(TenantUrlKey);
        }

        public static void Logout()
        {
            StorageService.DeleteData(TokenKey);
            StorageService.DeleteData(UserKey);
            StorageService.DeleteData(TenantUrlKey);
        }
    }
}
