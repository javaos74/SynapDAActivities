using System;
using System.Activities;
using System.Activities.Statements;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Charles.Synap.Activities
{
    public class SynapDAGetResult : CodeActivity
    {
        [Category("Login")]
        public InArgument<string> Endpoint { get; set; }

        [Category("Login")]

        public InArgument<string> ApiKey { get; set; }

        [Category("Input")]
        public InArgument<string> FID { get; set; }


        [Category("Input")]
        public InArgument<string> ZipFilePath { get; set; }

        [Category("Output")]
        public OutArgument<string> ErrorMessage { get; set; }

        [Category("Output")]
        public OutArgument<int> Status { get; set; }

        //[Category("Output")]
        //public OutArgument<string> FullText { get; set; }

        private UiPathHttpClient _httpClient;
        private String apikey;
        private String endpoint;
        private String fid;
        private string zipfilepath;

        public SynapDAGetResult()
        {
            // Constructor logic here
            _httpClient = new UiPathHttpClient();
        }   
  protected  override async void Execute(CodeActivityContext context)
        {
#if DEBUG
            //Debugger.Launch();
#endif
            SynapZipResponse _result;
            endpoint = Endpoint.Get(context);
            apikey = ApiKey.Get(context);
            fid = FID.Get(context);
            zipfilepath = ZipFilePath.Get(context);

            _httpClient.setEndpoint(endpoint);
            _httpClient.AddField("api_key", apikey);

            try
            {
                _result = await _httpClient.GetDAZipResult($"/result-all/{fid}", zipfilepath);
            }
            catch (Exception ex)
            {
                this.ErrorMessage.Set(context, ex.Message);
                this.Status.Set(context, (int)System.Net.HttpStatusCode.InternalServerError);
            }
        }
    }
}
