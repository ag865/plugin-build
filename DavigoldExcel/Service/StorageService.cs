using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace DavigoldExcel.Service
{
    internal class StorageService
    {
        public static void StoreData<T>(string key, T data)
        {
            string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\one\DavigoldAddin";
            string filePath = Path.Combine(folderPath, key + ".dat");

            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            byte[] encryptedToken = ProtectedData.Protect(Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(data)), null, DataProtectionScope.CurrentUser);

            File.WriteAllBytes(filePath, encryptedToken);
        }

        public static T GetStoredData<T>(string key)
        {
            string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\one\DavigoldAddin";
            string filePath = Path.Combine(folderPath, key + ".dat");

            if (File.Exists(filePath))
            {
                byte[] encryptedToken = File.ReadAllBytes(filePath);
                byte[] decryptedToken = ProtectedData.Unprotect(encryptedToken, null, DataProtectionScope.CurrentUser);
                return JsonConvert.DeserializeObject<T>(Encoding.UTF8.GetString(decryptedToken));
            }
            else
            {
                return default(T);
            }
        }

        public static void DeleteData(string key)
        {
            string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\one\DavigoldAddin";
            string filePath = Path.Combine(folderPath, key + ".dat");

            if (File.Exists(filePath))
            {
                
               File.Delete(filePath);
            }
        }
    }
}
