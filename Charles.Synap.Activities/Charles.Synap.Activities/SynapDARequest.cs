using System.Activities;
using System.ComponentModel;
using System.Diagnostics;
using System.Net;
using System.Runtime.CompilerServices;
using System.Text;
using Newtonsoft.Json.Linq;


namespace Charles.Synap.Activities
{
    public class SynapDARequest : AsyncCodeActivity // This base class exposes an OutArgument named Result
    {
        [Category("Login")]
        public InArgument<string> Endpoint { get; set; }

        [Category("Login")]

        public InArgument<string> ApiKey { get; set; }

        [Category("Input")]
        public InArgument<string> InputFilePath { get; set; }


        [Category("Output")]
        public OutArgument<string> FID { get; set; }

        //[Category("Output")]
        //public OutArgument<int> NumberOfPage { get; set; }

        [Category("Output")]
        public OutArgument<string> ErrorMessage { get; set; }

        [Category("Output")]
        public OutArgument<int> Status { get; set; }

        //[Category("Output")]
        //public OutArgument<string> FullText { get; set; }

        private UiPathHttpClient _httpClient;

        private string fid = string.Empty;
        private string errorMessage;
        private int status;

        public SynapDARequest()
        {
            // Constructor logic here
            _httpClient = new UiPathHttpClient();
        }
        /*
         * The returned value will be used to set the value of the Result argument
         */
        private async void Execute( string endpoint, string apikey, string filepath)
        {
            SynapDAResponse _result;

            _httpClient.setEndpoint(endpoint);
            _httpClient.AddFile(filepath, "file");
            _httpClient.AddField("type", "upload");
            _httpClient.AddField("api_key", apikey);

            _result = await _httpClient.UploadSynapDA( "/da");

            if (_result.status == HttpStatusCode.OK)
            {
                StringBuilder sb = new StringBuilder();
                JObject respJson = JObject.Parse(_result.body);

                status = (int)respJson["status"];
                fid = string.Empty;
                if (status == 200)
                {
                    fid = respJson["result"]["fid"].ToString();
                }
                else
                {
                    errorMessage =  respJson["result"].ToString();
                }
            }
            else
            {
                JObject respJson = JObject.Parse(_result.body);
                errorMessage = respJson["result"].ToString();
                status = (int)_result.status;
            }

        }

        protected override IAsyncResult BeginExecute(AsyncCodeActivityContext context, AsyncCallback callback, object state)
        {
            var apikey = ApiKey.Get(context);
            var endpoint = Endpoint.Get(context);   
            var filepath = InputFilePath.Get(context);

            var task = new Task(_ => Execute( endpoint, apikey, filepath), state);
            task.Start();
            if (callback != null)
            {
                task.ContinueWith(s => callback(s));
                task.Wait();
            }
            return task;
        }

        protected override void EndExecute(AsyncCodeActivityContext context, IAsyncResult result)
        {
            var task = (Task)result;

            if(task.IsCompletedSuccessfully)
            {
                ErrorMessage.Set(context, errorMessage);
                FID.Set(context, fid);
                Status.Set(context, status);
            }
            else
            {
                ErrorMessage.Set(context, "Error in SynapDARequest");
                FID.Set(context, string.Empty);
                Status.Set(context, 0);
            }
        }
    }
}
