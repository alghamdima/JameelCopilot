using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using JameelCopilot.Services;
using Xunit;

namespace JameelCopilot.Tests.Services
{
    public class ApiErrorsTests
    {
        [Fact]
        public void NotConfigured_Copy()
        {
            Assert.Equal("Jameel Copilot is not configured. Contact IT.", ApiErrors.NotConfigured);
        }

        [Theory]
        [InlineData(HttpStatusCode.Unauthorized)]
        [InlineData(HttpStatusCode.Forbidden)]
        public void SignInFailures(HttpStatusCode status)
        {
            Assert.Equal("Jameel Copilot could not sign in to the AI service. Contact IT.", ApiErrors.ForStatus(status, null));
        }

        [Fact]
        public void Busy_WithRetryAfter_ShowsTheSeconds()
        {
            Assert.Equal(
                "The AI service is busy. Try again in 12 seconds.",
                ApiErrors.ForStatus((HttpStatusCode)429, TimeSpan.FromSeconds(12)));
        }

        [Fact]
        public void Busy_RoundsUpAFractionOfASecond()
        {
            Assert.Equal(
                "The AI service is busy. Try again in 2 seconds.",
                ApiErrors.ForStatus((HttpStatusCode)429, TimeSpan.FromSeconds(1.2)));
        }

        [Theory]
        [InlineData(null)]
        [InlineData(0d)]
        [InlineData(-5d)]
        public void Busy_WithoutAUsableRetryAfter(double? seconds)
        {
            var retryAfter = seconds.HasValue ? TimeSpan.FromSeconds(seconds.Value) : (TimeSpan?)null;

            Assert.Equal("The AI service is busy. Try again shortly.", ApiErrors.ForStatus((HttpStatusCode)429, retryAfter));
        }

        [Theory]
        [InlineData(HttpStatusCode.InternalServerError, "The AI service returned an error (HTTP 500).")]
        [InlineData(HttpStatusCode.BadRequest, "The AI service returned an error (HTTP 400).")]
        [InlineData(HttpStatusCode.Redirect, "The AI service returned an error (HTTP 302).")]
        public void OtherStatuses_ShowTheCode(HttpStatusCode status, string expected)
        {
            Assert.Equal(expected, ApiErrors.ForStatus(status, null));
        }

        [Fact]
        public void NetworkFailures_CannotReach()
        {
            const string expected = "Jameel Copilot cannot reach the AI service.";

            Assert.Equal(expected, ApiErrors.ForException(new HttpRequestException("name not resolved")));
            Assert.Equal(expected, ApiErrors.ForException(new WebException("tls")));
            Assert.Equal(expected, ApiErrors.ForException(new TaskCanceledException()));
            Assert.Equal(expected, ApiErrors.ForException(new IOException("reset")));
        }
    }
}
