using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Net;
using System.Net.Http;
using System.Globalization;
using System.IO;
using System.Activities;
using System.Net.Sockets;
using UiPath.Platform.ResourceHandling;
using UiPath.Platform.ResourceHandling.Internals;
using System.IO.Abstractions;

namespace Charles.Synap.Activities
{
    public class ClovaResponse
    {
        public HttpStatusCode status { get; set; }
        public string body { get; set; }
    }
    public class UpstageResponse
    {
        public HttpStatusCode status { get; set; }
        public string body { get; set; }
    }
    public class SynapDAResponse
    {
        public HttpStatusCode status { get; set; }
        public string body { get; set; }
    }
    public class SynapZipResponse
    {
        public HttpStatusCode status { get; set; }
        public string filePath { get; set; }
    }
    internal class ClovaSpeechParamBoosting
    {
        public ClovaSpeechParamBoosting( string w) {
            this.words = w;
        }

        public string words { get; set; }
    }
    internal class ClovaSpeechParam
    {
        public string language { get; set; } = "ko-KR";
        public string completion { get; set; } = "sync";
        public bool wordAlignment { get; set; } = false;
        public bool fullText { get; set; } = true;
        public bool resultToObs { get; set; } = false;
        public bool noiseFiltering { get; set; } = true;
        public ClovaSpeechParamBoosting[] boostings { get; set; }
        public bool useDomainBoostings { get; set; } = false;
        public string forbiddens { get; set; }
    }
    public class UiPathHttpClient
    {

        public UiPathHttpClient() :
            this("https://ailab.synap.co.kr")
        {
        }
        public UiPathHttpClient( string endpoint)
        {
            this.url = endpoint;
            this.client = new HttpClient();
            this.client.Timeout = new TimeSpan(0, 3, 0);
            this.content = new MultipartFormDataContent("ocr----" + DateTime.Now.Ticks.ToString());
        }

        public void setEndpoint( string endpoint)
        {
            if (!string.IsNullOrEmpty(endpoint))
            {
                this.url = endpoint;
            }
        }
        public void setSecret( string secret)
        {
            setOCRSecret(secret);
        }
        public void setOCRSecret(string secret)
        {
            this.client.DefaultRequestHeaders.Add("X-OCR-SECRET", secret);
        }
        public void setSpeechSecret(string secret)
        {
            this.client.DefaultRequestHeaders.Add("X-CLOVASPEECH-API-KEY", secret);
        }
        public void setAuthorizationToken(string token)
        {
            this.client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
            this.client.DefaultRequestHeaders.Add("Accept", "*/*");
            //this.client.DefaultRequestHeaders.Add("User-Agent", "dotnet/1.0.0");
        }

        public async void AddFileResource(IResource file, string fieldName = "file")
        {
            var freader = file.GetReaderOrLocal();
            var fstream = await freader.OpenStreamAsync();
#if DEBUG
            Console.WriteLine($"file size: {file.GetSizeInBytes()}");
#endif
            byte[] buf = new byte[file.GetSizeInBytes()];
            int read_bytes = 0;
            int offset = 0;
            int remains = (int)buf.Length;
            do
            {
                read_bytes += fstream.Read(buf, offset, remains);
                offset += read_bytes;
                remains -= read_bytes;
            } while (remains > 0);
            freader.Dispose();

            this.content.Add(new StreamContent(new MemoryStream(buf)), fieldName, System.IO.Path.GetFileName( file.FullName));
        }

        public void AddFile(string fileName, string fieldName = "file")
        {
            var fstream = System.IO.File.OpenRead(fileName);
#if DEBUG
            Console.WriteLine($"file size: {fstream.Length}");
#endif
            byte[] buf = new byte[fstream.Length];
            int read_bytes = 0;
            int offset = 0;
            int remains = (int)fstream.Length;
            do
            {
                read_bytes += fstream.Read(buf, offset, remains);
                offset += read_bytes;
                remains -= read_bytes;
            } while (remains > 0);
            fstream.Close();

            this.content.Add(new StreamContent(new MemoryStream(buf)), fieldName, System.IO.Path.GetFileName(fileName));
        }
 
        public void AddField( string name, string value)
        {
            this.content.Add(new StringContent(value), name);
        }

        public void Clear()
        {
            this.content.Dispose();
            this.content = new MultipartFormDataContent("ocr----" + DateTime.Now.Ticks.ToString());
        }

        public async Task<ClovaResponse> Upload()
        {
#if DEBUG
            Console.WriteLine("http content count :" + this.content.Count());
#endif
            using (var message = this.client.PostAsync(this.url, this.content))
            {
                ClovaResponse resp = new ClovaResponse();
                resp.status = message.Result.StatusCode;
                resp.body = await message.Result.Content.ReadAsStringAsync();
                return resp;
            }
        }

        public async Task<UpstageResponse> UploadUpstage()
        {
#if DEBUG
            Console.WriteLine("http content count :" + this.content.Count());
#endif
            using (var message = this.client.PostAsync(this.url, this.content))
            {
                UpstageResponse resp = new UpstageResponse();
                resp.status = message.Result.StatusCode;
                resp.body = await message.Result.Content.ReadAsStringAsync();
                return resp;
            }
        }

        public async Task<SynapDAResponse> UploadSynapDA(string path)
        {
#if DEBUG
            Console.WriteLine("http content count :" + this.content.Count());
#endif
            using (var message = this.client.PostAsync(this.url + path, this.content))
            {
                SynapDAResponse resp = new SynapDAResponse();
                resp.status = message.Result.StatusCode;
                resp.body = await message.Result.Content.ReadAsStringAsync();
#if DEBUG
                Console.WriteLine($"http response status = {resp.status} body = {resp.body}");
#endif
                return resp;
            }
        }

        public async Task<SynapDAResponse> GetDAPageResult( string path, string body)
        {
#if DEBUG
            Console.WriteLine("http content count :" + this.content.Count());
#endif
            var _content = new StringContent(body);
            using (var message = this.client.PostAsync(this.url + path, _content))
            {
                SynapDAResponse resp = new SynapDAResponse();
                resp.status = message.Result.StatusCode;
                resp.body = await message.Result.Content.ReadAsStringAsync();
                return resp;
            }
        }

        public SynapZipResponse GetDAZipResult(string path, string filepath)
        {
#if DEBUG
            Console.WriteLine("http content count :" + this.content.Count());
#endif
            SynapZipResponse resp = new SynapZipResponse();
            using (var message = this.client.PostAsync(this.url + path, this.content).Result)
            {
                resp.status = message.StatusCode;
                if(resp.status == HttpStatusCode.OK)
                {
                    if (message.Content.Headers.Contains("Content-Disposition"))
                    {
                        var contentDisposition = message.Content.Headers.ContentDisposition;
                        if (contentDisposition != null && contentDisposition.FileName != null)
                        {
                            string fileName = contentDisposition.FileName.Trim('"');
#if DEBUG
                            Console.WriteLine($"Filename: {fileName}");
#endif
                            using (var stream = message.Content.ReadAsStream()) 
                            {
                                resp.filePath = filepath; //Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", fileName);
                                using var fileStream = File.Create(filepath);
                                stream.CopyTo(fileStream);
                            }
#if DEBUG
                            Console.WriteLine($"File saved to: {filepath}");
#endif
                        }
                    }
                    else
                    {
                        resp.status = HttpStatusCode.NotFound;
                        resp.filePath = string.Empty;
                    }

                }
                else
                {
                    resp.filePath = string.Empty;
                }
            }
#if DEBUG
            Console.WriteLine($"http response status = {resp.status} filePath = {resp.filePath}");  
#endif
            return resp;

        }

        private HttpClient client;
        private string url;
        private MultipartFormDataContent content;
    }
}
