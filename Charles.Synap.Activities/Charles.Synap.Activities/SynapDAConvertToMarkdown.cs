using OfficeOpenXml;
using System;
using System.Activities;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using UiPath.Platform.ResourceHandling;
using UiPath.Platform.ResourceHandling.Internals;

namespace Charles.Synap.Activities
{
    public class SynapDAConvertToMarkdown : CodeActivity
    {
        public InArgument<IResource> ResultZip { get; set; }
        public OutArgument<int> PageCount { get; set; }
        public OutArgument<string> ErrorMessage { get; set; }
        public OutArgument<string> MarkdownBody { get; set; }

        private string mdbody; // Output file path
        private int pageCount; // Page count for the markdown file
        public SynapDAConvertToMarkdown()
        {
            // Constructor logic can be added here if needed
        }

        private int MergeAllPageIntoMarkdown(IResource  zipFile)
        {
            StringBuilder _mdbodybuffer = new StringBuilder();
            pageCount = 0;
            ZipArchive archive = null;
            try
            {
                string tempRoot = Path.GetTempPath();
                string uniqueDirectoryName = Guid.NewGuid().ToString();
                string tempDirectoryPath = Path.Combine(tempRoot, uniqueDirectoryName);
                //var fstream = zipFile.GetReaderOrLocal().OpenStreamAsync().Result;
                //using (FileStream fileStream = new FileStream(fstream, FileMode.Open, FileAccess.Read))
                using( Stream fileStream = zipFile.GetReaderOrLocal().OpenStreamAsync().Result) 
                {
                    using (archive = new ZipArchive(fileStream, ZipArchiveMode.Read))
                    {
                        // 압축 파일 내에서 확장자가 ".xml"인 파일들을 필터링합니다.
                        var mdFiles = archive.Entries.Where(entry => entry.FullName.EndsWith(".md", StringComparison.OrdinalIgnoreCase))
                                                     .OrderBy(entry => entry.Name, StringComparer.OrdinalIgnoreCase);
#if DEBUG
                        Console.WriteLine($"압축 파일 '{zipFile.FullName}'에서 MD 파일 추출 시작...");
#endif

                        foreach (ZipArchiveEntry entry in mdFiles)
                        {
                            string extractedFilePath = Path.Combine(tempDirectoryPath, entry.FullName);

                            // 디렉토리 구조를 유지하기 위해 필요한 하위 디렉토리를 생성합니다.
                            string directoryName = Path.GetDirectoryName(extractedFilePath);
                            if (!string.IsNullOrEmpty(directoryName) && !Directory.Exists(directoryName))
                            {
                                Directory.CreateDirectory(directoryName);
                            }
#if DEBUG
                            Console.WriteLine($"  '{entry.FullName}' 추출 중...");
#endif
                            entry.ExtractToFile(extractedFilePath, true); // true: 이미 파일이 존재하면 덮어쓰기

#if DEBUG
                            Console.WriteLine($"  '{entry.FullName}' 추출 & markdown merge 완료: '{extractedFilePath}'");
#endif
                            _mdbodybuffer.AppendLine(File.ReadAllText(extractedFilePath)); // Read the content of the extracted file and append it to mdbody
                            _mdbodybuffer.AppendLine("\\newpage"); // Add a new line for separation
                            File.Delete(extractedFilePath); // 변환 후 XML 파일 삭제
                            pageCount++;
                        }
                    }
                    Directory.Delete(tempDirectoryPath, true); // 변환 후 임시 디렉토리 삭제
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"오류 발생: {ex.Message}");
                mdbody = string.Empty; // Set the output to empty in case of error
                pageCount = -1; // Set the page count to -1 in case of error
                if(archive != null)
                {
                    archive.Dispose(); // Dispose the archive if it was opened
                }   
            }
            mdbody = _mdbodybuffer.ToString(); // Assign the merged content to mdbody
            return pageCount; // Return the number of pages converted
        }

        /*
        protected override IAsyncResult BeginExecute(AsyncCodeActivityContext context, AsyncCallback callback, object state)
        {
            var zipfile = ResultZipFIle.Get(context);

            var task = new Task(_ => MergeAllPageIntoMarkdown(zipfile), state);
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

            if (task.IsCompletedSuccessfully)
            {
                PageCount.Set(context, pageCount);
                ErrorMessage.Set(context, string.Empty);
                MarkdownBody.Set(context, mdbody); // Set the merged markdown content to the output argument
            }
            else
            {
                ErrorMessage.Set(context, "Error in SynapDARequest");
                PageCount.Set(context, -1);
                MarkdownBody.Set(context, string.Empty); // Set the output to empty in case of error
            }
        }
        */
        protected override void Execute(CodeActivityContext context)
        {
            var zipfile = ResultZip.Get(context);
            if( MergeAllPageIntoMarkdown(zipfile) > 0)
            {
                PageCount.Set(context, pageCount);
                ErrorMessage.Set(context, string.Empty);
                MarkdownBody.Set(context, mdbody); // Set the merged markdown content to the output argument
            }
            else
            {
                ErrorMessage.Set(context, "Error in SynapDAConvertToMarkdown");
                PageCount.Set(context, -1);
                MarkdownBody.Set(context, string.Empty); // Set the output to empty in case of error
            }
        }
    }
}
