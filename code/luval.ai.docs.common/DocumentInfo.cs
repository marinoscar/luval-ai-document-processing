using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace luval.ai.docs.common
{
    public class DocumentInfo
    {
        public DocumentInfo()
        {

        }

        public DocumentInfo(string localFilePath)
        {
            Uri = new Uri(localFilePath);
        }

        public string DocumentType { get; set; }
        public Uri Uri { get; set; }

        public async Task<Stream> GetContentAsync()
        {
            return Uri.IsFile ? GetFileContent() : await GetWebContentAsync();
        }

        protected virtual async Task<Stream> GetWebContentAsync()
        {
            var client = new HttpClient();
            var response = await client.GetAsync(Uri);
            if (!response.IsSuccessStatusCode) throw new HttpRequestException($"Failed to download file { Uri } with status { response.StatusCode } and message { response.RequestMessage }");
            var stream = new MemoryStream();
            await response.Content.CopyToAsync(stream);
            return stream;
        }

        protected virtual Stream GetFileContent()
        {
            var file = new FileInfo(Uri.LocalPath);
            if (file.Exists) throw new ArgumentException($"File { file.FullName } doesn't exists");
            var stream = new StreamReader(file.FullName);
            return stream.BaseStream;
        }
    }
}
