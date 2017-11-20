using System;
using System.Collections.Concurrent;
using System.IO;
using System.Net.Http;
using System.Threading;

namespace SpPrefetchIndexBuilder
{
    class FileDownloader
    {
        private BlockingCollection<FileToDownload> fileDownloadBlockingCollection;
        private HttpClient client;
        private System.Collections.Generic.Dictionary<int, HttpClient> webClients = new System.Collections.Generic.Dictionary<int, HttpClient>();

        public FileDownloader(BlockingCollection<FileToDownload> fileDownloadBlockingCollection, HttpClient client)
        {
            this.client = client;
            this.fileDownloadBlockingCollection = fileDownloadBlockingCollection;
        }

        public void StartDownloads(int timeout)
        {
            try 
            {
				Console.WriteLine("Starting Thread {0}", Thread.CurrentThread.ManagedThreadId);
				FileToDownload toDownload;
				while (fileDownloadBlockingCollection.TryTake(out toDownload))
				{
                    SpPrefetchIndexBuilder.CheckAbort();
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
                        Console.WriteLine("Got error trying to download url {0} to file {1}: {2}", toDownload.site + toDownload.serverRelativeUrl, toDownload.saveToPath, e.Message);
						Console.WriteLine(e.StackTrace);
					}
					Console.WriteLine("Thread {0} - Finished attempt to download {1} to {2}", Thread.CurrentThread.ManagedThreadId, toDownload.serverRelativeUrl, toDownload.saveToPath);    
                }
            }
            catch (Exception e2) 
            {
                Console.WriteLine("Thread {0} File Downloader failed - {1}", Thread.CurrentThread.ManagedThreadId, e2);
				Console.WriteLine(e2.StackTrace);
            }
        }

        public static void DownloadFiles(BlockingCollection<FileToDownload> fileDownloadBlockingCollection, int timeoutInMilliSec, HttpClient client)
        {
            new FileDownloader(fileDownloadBlockingCollection, client).StartDownloads(timeoutInMilliSec);
        }
    }
}