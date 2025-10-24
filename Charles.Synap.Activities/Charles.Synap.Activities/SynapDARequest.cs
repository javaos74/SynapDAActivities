using System.Activities;
using System.ComponentModel;
using System.Diagnostics;
using System.Net;
using System.Resources;
using System.Runtime.CompilerServices;
using System.Text;
using Newtonsoft.Json.Linq;
using UiPath.Platform.ResourceHandling;


namespace Charles.Synap.Activities
{
    public class SynapDARequest : CodeActivity 
    {
        [Category("Login")]
        public InArgument<string> Endpoint { get; set; }

        [Category("Login")]

        public InArgument<string> ApiKey { get; set; }

        [Category("Input")]
        public InArgument<IResource> InputFile { get; set; }


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

        protected override void Execute(CodeActivityContext context)
        {
            var apikey = ApiKey.Get(context);
            var endpoint = Endpoint.Get(context);
            var fileresource = InputFile.Get(context);

            _httpClient.Clear();
            try
            {
                var resp = _Execute(endpoint, apikey, fileresource);
                ErrorMessage.Set(context, string.Empty);
                FID.Set(context, fid);
                Status.Set(context, status);
            }
            catch (Exception ex)
            {
                ErrorMessage.Set(context, ex.Message);
                Status.Set(context, (int)System.Net.HttpStatusCode.InternalServerError);
                FID.Set(context, string.Empty);
            }
        }
        /*
         * The returned value will be used to set the value of the Result argument
         */
        private SynapDAResponse _Execute( string endpoint, string apikey, IResource fileresource)
        {
            SynapDAResponse _result;

            _httpClient.setEndpoint(endpoint);
            _httpClient.AddFileResource(fileresource, "file");
            _httpClient.AddField("type", "upload");
            _httpClient.AddField("api_key", apikey);

            _result =  _httpClient.UploadSynapDA( "/da").Result;

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
                errorMessage = _result.body;
                status = (int)_result.status;
            }
            return _result;

        }
        /*

        protected override IAsyncResult BeginExecute(AsyncCodeActivityContext context, AsyncCallback callback, object state)
        {
            var apikey = ApiKey.Get(context);
            var endpoint = Endpoint.Get(context);   
            var fileresource = InputFile.Get(context);

            var task = new Task(_ => _Execute( endpoint, apikey, fileresource), state);
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
                ErrorMessage.Set(context, string.Empty);
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
        */
    }
}
