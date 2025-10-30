using System;
using System.Activities;
using System.ComponentModel;
using System.Diagnostics;
using System.Net;
using System.Text;
using Newtonsoft.Json.Linq;

namespace Charles.Synap.Activities
{
    public class SynapDACheckFileStatus : CodeActivity
    {
        [Category("Login")]
        public InArgument<string> Endpoint { get; set; }

        [Category("Login")]
        public InArgument<string> ApiKey { get; set; }

        [Category("Input")]
        public InArgument<string> FID { get; set; }

        [Category("Output")]
        public OutArgument<string> FileStatus { get; set; }

        [Category("Output")]
        public OutArgument<int> TotalPages { get; set; }

        [Category("Output")]
        public OutArgument<int> ReturnCode { get; set; }

        [Category("Output")]
        public OutArgument<int> Status { get; set; }

        [Category("Output")]
        public OutArgument<string> ErrorMessage { get; set; }

        private UiPathHttpClient _httpClient;
        private string fileStatus = string.Empty;
        private int totalPages = 0;
        private int returnCode = 0;
        private int status = 0;
        private string errorMessage = string.Empty;

        public SynapDACheckFileStatus()
        {
            _httpClient = new UiPathHttpClient();
        }

        protected override void Execute(CodeActivityContext context)
        {
            var apikey = ApiKey.Get(context);
            var endpoint = Endpoint.Get(context);
            var fid = FID.Get(context);

            _httpClient.Clear();
            try
            {
                var resp = _Execute(endpoint, apikey, fid);
                
                FileStatus.Set(context, fileStatus);
                TotalPages.Set(context, totalPages);
                ReturnCode.Set(context, returnCode);
                Status.Set(context, status);
                ErrorMessage.Set(context, errorMessage);
            }
            catch (Exception ex)
            {
                FileStatus.Set(context, string.Empty);
                TotalPages.Set(context, 0);
                ReturnCode.Set(context, 0);
                Status.Set(context, (int)HttpStatusCode.InternalServerError);
                ErrorMessage.Set(context, ex.Message);
            }
        }

        private SynapDAResponse _Execute(string endpoint, string apikey, string fid)
        {
#if DEBUG
            //Debugger.Launch();
#endif
            SynapDAResponse _result;

            _httpClient.setEndpoint(endpoint);
            
            // JSON 바디 생성
            var requestBody = new JObject
            {
                ["api_key"] = apikey
            };

            string jsonBody = requestBody.ToString();
            string path = $"/filestatus/{fid}";

            _result = _httpClient.GetDAFileStatus(path, jsonBody).Result;

            if (_result.status == HttpStatusCode.OK)
            {
                JObject respJson = JObject.Parse(_result.body);
                
                status = (int)respJson["status"];
                
                if (status == 200)
                {
                    var result = respJson["result"];
                    fileStatus = result["filestatus"]?.ToString() ?? string.Empty;
                    
                    // total_pages는 LOADING을 제외한 상태에서만 제공됨
                    if (fileStatus != "LOADING" && result["total_pages"] != null)
                    {
                        totalPages = (int)result["total_pages"];
                    }
                    else
                    {
                        totalPages = 0;
                    }
                    
                    // returncode는 SUCCESS와 FAILED 상태에서만 제공됨
                    if ((fileStatus == "SUCCESS" || fileStatus == "FAILED") && result["returncode"] != null)
                    {
                        returnCode = (int)result["returncode"];
                    }
                    else
                    {
                        returnCode = 0;
                    }
                    
                    errorMessage = string.Empty;
                }
                else
                {
                    fileStatus = string.Empty;
                    totalPages = 0;
                    returnCode = 0;
                    errorMessage = respJson["result"]?.ToString() ?? "Unknown error";
                }
            }
            else
            {
                fileStatus = string.Empty;
                totalPages = 0;
                returnCode = 0;
                errorMessage = _result.body;
                status = (int)_result.status;
            }

#if DEBUG
            Console.WriteLine($"FileStatus: {fileStatus}, TotalPages: {totalPages}, ReturnCode: {returnCode}, Status: {status}");
#endif

            return _result;
        }
    }
}