using System;
using System.Collections.Concurrent;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;

namespace SpPrefetchIndexBuilder
{
    class FileDownloader
    {
        private BlockingCollection<FileToDownload> fileDownloadBlockingCollection;
        private System.Net.CredentialCache cc;
        private System.Collections.Generic.Dictionary<int, HttpClient> webClients = new System.Collections.Generic.Dictionary<int, HttpClient>();

        public FileDownloader(BlockingCollection<FileToDownload> fileDownloadBlockingCollection, System.Net.CredentialCache cc)
        {
            this.cc = cc;
            this.fileDownloadBlockingCollection = fileDownloadBlockingCollection;
        }

        public void StartDownloads(int timeout)
        {
            Console.WriteLine("Starting Thread {0}", Thread.CurrentThread.ManagedThreadId);
            FileToDownload toDownload;
			HttpClientHandler handler = new HttpClientHandler();
			handler.Credentials = cc;
			HttpClient client = new HttpClient(handler);
            while (fileDownloadBlockingCollection.TryTake(out toDownload))
            {
                Console.WriteLine("Thread {0} - Starting download of {1} to {2}", Thread.CurrentThread.ManagedThreadId, toDownload.serverRelativeUrl, toDownload.saveToPath);
				try
                {
                    var responseResult = client.GetAsync(toDownload.site + toDownload.serverRelativeUrl);
					using (var memStream = responseResult.Result.Content.ReadAsStreamAsync().Result)
					{
						using (var fileStream = File.Create(toDownload.saveToPath))
						{
							memStream.CopyTo(fileStream);
						}

					}
                }
				catch (Exception e)
				{
					Console.WriteLine("Got error trying to download file {0}: {1}", toDownload.saveToPath, e.Message);
					Console.WriteLine(e.StackTrace);
				}
                Console.WriteLine("Thread {0} - Finished attempt to download {1} to {2} - Success? {3}", Thread.CurrentThread.ManagedThreadId, toDownload.serverRelativeUrl, toDownload.saveToPath);
            }
        }

        public static void DownloadFiles(BlockingCollection<FileToDownload> fileDownloadBlockingCollection, int timeoutInMilliSec, System.Net.CredentialCache cc)
        {
            new FileDownloader(fileDownloadBlockingCollection, cc).StartDownloads(timeoutInMilliSec);
        }
    }
}