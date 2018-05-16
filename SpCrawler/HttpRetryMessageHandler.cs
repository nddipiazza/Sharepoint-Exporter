using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Polly;

namespace SpPrefetchIndexBuilder {
  public class HttpRetryMessageHandler : DelegatingHandler {
    private int numRetries;

    public HttpRetryMessageHandler(HttpClientHandler handler, int numRetries) : base(handler) {
      this.numRetries = numRetries;
    }

    protected override Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request,
        CancellationToken cancellationToken) =>
        Policy
            .Handle<HttpRequestException>()
            .Or<TaskCanceledException>()
            .OrResult<HttpResponseMessage>(x => !x.IsSuccessStatusCode)
            .WaitAndRetryAsync(numRetries, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)))
            .ExecuteAsync(() => base.SendAsync(request, cancellationToken));
  }
}
