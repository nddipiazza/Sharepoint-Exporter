using System;
using System.Collections.Concurrent;
using System.IO;
using System.Net.Http;
using System.Threading;

namespace SpPrefetchIndexBuilder
{
    class FileDownloader
    {
        static int NUM_RETRIES = 3;

        BlockingCollection<FileToDownload> fileDownloadBlockingCollection;
        HttpClient client;
        System.Collections.Generic.Dictionary<int, HttpClient> webClients = new System.Collections.Generic.Dictionary<int, HttpClient>();

        public FileDownloader(BlockingCollection<FileToDownload> fileDownloadBlockingCollection, HttpClient client)
        {
            this.fileDownloadBlockingCollection = fileDownloadBlockingCollection;
            this.client = client;
        }

        public void AttemptToDownload(FileToDownload toDownload, int numRetry)
        {
            try
            {
                var responseResult = client.GetAsync(SpPrefetchIndexBuilder.topParentSite + toDownload.serverRelativeUrl);
                if (responseResult.Result != null && responseResult.Result.StatusCode == System.Net.HttpStatusCode.OK)
                {
                    using (var memStream = responseResult.Result.Content.ReadAsStreamAsync().GetAwaiter().GetResult())
                    {
                        using (var fileStream = File.Create(toDownload.saveToPath))
                        {
                            memStream.CopyTo(fileStream);
                        }
                    }
                    Console.WriteLine("Thread {0} - Successfully downloaded {1} to {2}", Thread.CurrentThread.ManagedThreadId, toDownload.serverRelativeUrl, toDownload.saveToPath);
                }
                else
                {
                    Console.WriteLine("Got non-OK status {0} when trying to download url {1}", responseResult.Result.StatusCode, SpPrefetchIndexBuilder.topParentSite + toDownload.serverRelativeUrl);
                }
            }
            catch (Exception e)
            {
                if (numRetry >= NUM_RETRIES)
                {
                    Console.WriteLine("Gave up trying to download url {0} to file {1} after {2} retries due to error: {3}", SpPrefetchIndexBuilder.topParentSite + toDownload.serverRelativeUrl, toDownload.saveToPath, NUM_RETRIES, e);
                }
                else
                {
                    AttemptToDownload(toDownload, numRetry + 1);
                }
            }
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
                    AttemptToDownload(toDownload, 0);

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