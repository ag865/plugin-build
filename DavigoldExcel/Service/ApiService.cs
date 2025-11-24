using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace DavigoldExcel.Service
{
    public class ApiService
    {
        public string Token { set; get; }


        private readonly HttpClient _httpClient;

        public ApiService() {
            // Configure security protocol globally
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
            
            // For .NET Framework 4.7+, also support TLS 1.3 if available
            try
            {
                ServicePointManager.SecurityProtocol |= (SecurityProtocolType)3072; // TLS 1.3
            }
            catch { }
            
            // Add certificate validation callback to handle SSL issues
            ServicePointManager.ServerCertificateValidationCallback += 
                (sender, certificate, chain, sslPolicyErrors) => true;
            
            // Create HttpClientHandler with SSL configuration
            var handler = new HttpClientHandler()
            {
                ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true,
                SslProtocols = System.Security.Authentication.SslProtocols.Tls12 | 
                               System.Security.Authentication.SslProtocols.Tls11 | 
                               System.Security.Authentication.SslProtocols.Tls
            };
            
            try
            {
                // Try to add TLS 1.3 support if available
                handler.SslProtocols |= (System.Security.Authentication.SslProtocols)12288; // Tls13
            }
            catch { }
            
            _httpClient = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromMinutes(30)
            };
        }

        private string ExtractErrorMessage(string errorResponse)
        {
            if (string.IsNullOrWhiteSpace(errorResponse))
            {
                return null;
            }

            try
            {
                // Try to parse as JSON and extract common error message fields
                var jsonObject = JObject.Parse(errorResponse);
                
                // Try common error message field names
                var errorMessage = jsonObject["error"]?.ToString() 
                    ?? jsonObject["message"]?.ToString() 
                    ?? jsonObject["errorMessage"]?.ToString()
                    ?? jsonObject["Error"]?.ToString()
                    ?? jsonObject["Message"]?.ToString();
                
                if (!string.IsNullOrWhiteSpace(errorMessage))
                {
                    return errorMessage;
                }
            }
            catch
            {
                // If parsing fails, return the raw response
            }

            // If not JSON or no error field found, return the raw response
            return errorResponse;
        }

        private string GetExceptionMessage(Exception ex)
        {
            if (ex == null)
            {
                return "An error occurred while making the request.";
            }

            // First, check if there's an inner exception with a more specific message
            // Inner exceptions often contain the actual network/system error
            if (ex.InnerException != null && !string.IsNullOrWhiteSpace(ex.InnerException.Message))
            {
                // Prefer inner exception message as it usually has the real error
                // (e.g., "No connection could be made because the target machine actively refused it")
                return ex.InnerException.Message;
            }

            // If no inner exception or it's empty, use the exception's own message
            if (!string.IsNullOrWhiteSpace(ex.Message))
            {
                return ex.Message;
            }

            // Fallback message
            return "An error occurred while making the request.";
        }

        private bool IsCloudflareError(HttpResponseMessage response, string errorResponse)
        {
            if (response == null)
                return false;

            // Check for Cloudflare-specific headers
            if (response.Headers.Contains("CF-RAY") || 
                response.Headers.Contains("cf-mitigated"))
            {
                return true;
            }

            // Check Server header
            if (response.Headers.Contains("Server"))
            {
                var serverHeader = response.Headers.GetValues("Server").FirstOrDefault();
                if (serverHeader != null && serverHeader.ToLower().Contains("cloudflare"))
                {
                    return true;
                }
            }

            // Check response content for Cloudflare indicators
            if (!string.IsNullOrWhiteSpace(errorResponse))
            {
                string lowerResponse = errorResponse.ToLower();
                if (lowerResponse.Contains("cloudflare") || 
                    lowerResponse.Contains("cf-ray") ||
                    lowerResponse.Contains("checking your browser") ||
                    lowerResponse.Contains("just a moment") ||
                    lowerResponse.Contains("ddos protection") ||
                    lowerResponse.Contains("access denied"))
                {
                    return true;
                }
            }

            return false;
        }

        private bool IsBotProtectionError(HttpResponseMessage response, string errorResponse)
        {
            if (!IsCloudflareError(response, errorResponse))
                return false;

            // Check for bot protection specific indicators
            if (response.Headers.Contains("cf-mitigated"))
            {
                var cfMitigated = response.Headers.GetValues("cf-mitigated").FirstOrDefault();
                if (cfMitigated != null && cfMitigated.ToLower().Contains("challenge"))
                {
                    return true;
                }
            }

            // Check response content for bot protection messages
            if (!string.IsNullOrWhiteSpace(errorResponse))
            {
                string lowerResponse = errorResponse.ToLower();
                if (lowerResponse.Contains("checking your browser") ||
                    lowerResponse.Contains("just a moment") ||
                    lowerResponse.Contains("bot") ||
                    lowerResponse.Contains("challenge") ||
                    lowerResponse.Contains("captcha"))
                {
                    return true;
                }
            }

            // 403 with Cloudflare usually means bot protection
            if (response.StatusCode == HttpStatusCode.Forbidden && IsCloudflareError(response, errorResponse))
            {
                return true;
            }

            return false;
        }

        public async Task<T> GetJsonResponse<T>(string apiUrl, string authToken)
        {
            var request = new HttpRequestMessage(HttpMethod.Get, apiUrl);
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authToken);

            HttpResponseMessage response = await _httpClient.SendAsync(request);

            if (response.IsSuccessStatusCode)
            {
                string jsonResponse = await response.Content.ReadAsStringAsync();
                return Newtonsoft.Json.JsonConvert.DeserializeObject<T>(jsonResponse);
            }
            else
            {
                // Handle unsuccessful response (e.g., log error, throw exception)
                return default(T);
            }
        }

        public async Task<TResponse> PostJsonData<TResponse>(string apiUrl, object requestData)
        {
            try
            {

                apiUrl = apiUrl.Replace(":80", ":443");
                apiUrl = apiUrl.Replace("http://", "https://");

                // Serialize the request data to JSON
                string jsonData = Newtonsoft.Json.JsonConvert.SerializeObject(requestData);

                var token = requestData.GetType().GetProperty("token")?.GetValue(requestData)?.ToString();

                // Create the request message
                var request = new HttpRequestMessage(HttpMethod.Post, apiUrl);
                request.Content = new StringContent(jsonData, Encoding.UTF8, "application/json");
                if (!String.IsNullOrWhiteSpace(Token))
                {
                    request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", Token);
                }

                else if (!string.IsNullOrWhiteSpace(token))
                {
                    request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                }

                const string yTokenValue = "1ae9724a-98f1-4c13-be95-f731dfa225c9";
                if (request.Headers.Contains("Y-Token")) request.Headers.Remove("Y-Token");
                request.Headers.TryAddWithoutValidation("Y-Token", yTokenValue);

                // Send the request
                HttpResponseMessage response = await _httpClient.SendAsync(request);

                if (response.IsSuccessStatusCode)
                {
                    string jsonResponse = await response.Content.ReadAsStringAsync();
                    return Newtonsoft.Json.JsonConvert.DeserializeObject<TResponse>(jsonResponse);
                } 
                else if (response.StatusCode == HttpStatusCode.Unauthorized)
                {
                    string errorResponse = await response.Content.ReadAsStringAsync();
                    string errorMessage = ExtractErrorMessage(errorResponse);
                    throw new HttpRequestException(errorMessage);
                }
                else
                {
                    // Handle unsuccessful response - extract error message from server
                    string errorResponse = await response.Content.ReadAsStringAsync();
                    
                    // Check for Cloudflare errors
                    if (IsBotProtectionError(response, errorResponse))
                    {
                        throw new HttpRequestException("Access blocked by Cloudflare bot protection. Please try again later or contact support.");
                    }
                    else if (IsCloudflareError(response, errorResponse))
                    {
                        throw new HttpRequestException("Access blocked by Cloudflare");
                    }
                    
                    // For other errors, try to extract the error message
                    string errorMessage = ExtractErrorMessage(errorResponse);
                    if (string.IsNullOrWhiteSpace(errorMessage))
                    {
                        errorMessage = $"Server returned status code: {response.StatusCode}";
                    }
                    throw new HttpRequestException(errorMessage);
                }
            } catch(HttpRequestException ex)
            {
                // Extract the actual error message from the exception chain
                string errorMessage = GetExceptionMessage(ex);
                if(errorMessage.ToLower() == "unable to connect to the remote server")
                {
                    errorMessage = "Sorry, this URL path does not exis!";
                }
                throw new HttpRequestException(errorMessage, ex);
            }
            catch (Exception ex) { 
                // Extract the actual error message from the exception chain
                string errorMessage = GetExceptionMessage(ex);
                throw new HttpRequestException(errorMessage, ex);
            }
        }

        public async Task<bool> SyncFileExcel(string apiUrl, object requestData, byte[] workbookData)
        {
            // Serialize the request data to JSON
            string jsonData = Newtonsoft.Json.JsonConvert.SerializeObject(requestData);

            // Create the request message
            var request = new HttpRequestMessage(HttpMethod.Post, apiUrl);
            StringContent jsonContent = new StringContent(jsonData, Encoding.UTF8, "application/json");
            ByteArrayContent fileContent = new ByteArrayContent(workbookData);

            if (!String.IsNullOrWhiteSpace(Token))
            {
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", Token);
            }

            // Create MultipartFormDataContent to combine both JSON and file data
            using (var formData = new MultipartFormDataContent())
            {
                formData.Add(jsonContent, "data");
                formData.Add(fileContent, "file", "newFile.xlsx");

                // Send HTTP POST request to the server
                //HttpResponseMessage response = await httpClient.PostAsync(apiUrl, formData);
                request.Content = formData;

                // Send the request
                HttpResponseMessage response = await _httpClient.SendAsync(request);

                if (response.IsSuccessStatusCode)
                {
                    return true;
                }
                else
                {
                    // Handle unsuccessful response (e.g., log error, throw exception)
                    return false;
                }

                // Handle the response as before
            }
        }

        public async Task<bool> SyncFilePresentation(string apiUrl, object requestData, byte[] workbookData)
        {
            // Serialize the request data to JSON
            string jsonData = Newtonsoft.Json.JsonConvert.SerializeObject(requestData);

            // Create the request message
            var request = new HttpRequestMessage(HttpMethod.Post, apiUrl);
            StringContent jsonContent = new StringContent(jsonData, Encoding.UTF8, "application/json");
            ByteArrayContent fileContent = new ByteArrayContent(workbookData);

            if (!String.IsNullOrWhiteSpace(Token))
            {
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", Token);
            }

            // Create MultipartFormDataContent to combine both JSON and file data
            using (var formData = new MultipartFormDataContent())
            {
                formData.Add(jsonContent, "data");
                formData.Add(fileContent, "file", "newFile.pptx");

                // Send HTTP POST request to the server
                //HttpResponseMessage response = await httpClient.PostAsync(apiUrl, formData);
                request.Content = formData;

                // Send the request
                HttpResponseMessage response = await _httpClient.SendAsync(request);

                if (response.IsSuccessStatusCode)
                {
                    return true;
                }
                else
                {
                    // Handle unsuccessful response (e.g., log error, throw exception)
                    return false;
                }

                // Handle the response as before
            }
        }

        public async Task<bool> SyncFileWord(string apiUrl, object requestData, byte[] workbookData)
        {
            // Serialize the request data to JSON
            string jsonData = Newtonsoft.Json.JsonConvert.SerializeObject(requestData);

            // Create the request message
            var request = new HttpRequestMessage(HttpMethod.Post, apiUrl);
            StringContent jsonContent = new StringContent(jsonData, Encoding.UTF8, "application/json");
            ByteArrayContent fileContent = new ByteArrayContent(workbookData);

            if (!String.IsNullOrWhiteSpace(Token))
            {
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", Token);
            }

            // Create MultipartFormDataContent to combine both JSON and file data
            using (var formData = new MultipartFormDataContent())
            {
                formData.Add(jsonContent, "data");
                formData.Add(fileContent, "file", "newFile.docx");

                // Send HTTP POST request to the server
                //HttpResponseMessage response = await httpClient.PostAsync(apiUrl, formData);
                request.Content = formData;

                // Send the request
                HttpResponseMessage response = await _httpClient.SendAsync(request);

                if (response.IsSuccessStatusCode)
                {
                    return true;
                }
                else
                {
                    // Handle unsuccessful response (e.g., log error, throw exception)
                    return false;
                }

                // Handle the response as before
            }
        }
    }
}
