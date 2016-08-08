//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using ExcelAsACalculationService.Models;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Web;
using System.Web.Mvc;
using System.Configuration;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using System.Text;
using Newtonsoft.Json.Linq;

namespace ExcelFunctionaaService.Controllers
{
    [Authorize]
    public class ExcelFunctionController : Controller
    {
        private string clientId = ConfigurationManager.AppSettings["ida:ClientId"];
        private string appKey = ConfigurationManager.AppSettings["ida:ClientSecret"];
        private string aadInstance = ConfigurationManager.AppSettings["ida:AADInstance"];
        private string graphResourceUrl = "https://graph.microsoft.com";
        private string accessToken = string.Empty;
        private string sessionId = string.Empty;

        // GET: ExcelFunction
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<ActionResult> Calculate(PMT pmt)
        {
            accessToken = await GetGraphAccessTokenAsync();

            // create a new workbook session
            var restUrl = string.Format("{0}/v1.0/me/drive/root:/book.xlsx:/workbook/createSession", graphResourceUrl);

            var sessionRequest = new SessionRequest { persistChanges = "true" };

            var result = await DoRequest(HttpMethod.Post, sessionRequest, restUrl);

            // get the sessionId from the result
            dynamic d = JObject.Parse(result);
            sessionId = d.id;

            // get the values from the modelstate class (pmt)

            // as a percentage per month
            decimal rate = pmt.InterestRate / 12/ 100;

            // inverse the amount (simply to show the value as a positive amount)
            decimal loanAmount = pmt.LoanAmount * -1;

            // number of months
            int numberOfMonths = pmt.NumberOfMonths;

            // call one of the built-in workbook functions (pmt)
            restUrl = string.Format(@"{0}/v1.0/me/drive/root:/book.xlsx:/workbook/functions/pmt", graphResourceUrl);
            var pmtRequest = new PMTRequest { Rate = rate, Nper = numberOfMonths, Pv = loanAmount };
            result = await DoRequest(HttpMethod.Post, pmtRequest, restUrl);

            // refresh the model's MonthlyPaymentAmount
            ModelState.Remove("MonthlyPaymentAmount");

            dynamic v = JObject.Parse(result);
            decimal value = v.value;

            // for financial purposes round the value to 2 dp
            pmt.MonthlyPaymentAmount = Math.Round(value, 2);

            return View("Index", pmt);
        }

        private async Task<string> GetGraphAccessTokenAsync()
        {
            var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
            var userObjectId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
            string tenantID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;

            var clientCredential = new ClientCredential(clientId, appKey);
            var userIdentifier = new UserIdentifier(userObjectId, UserIdentifierType.UniqueId);

            // create auth context
            AuthenticationContext authContext = new AuthenticationContext(aadInstance + tenantID, new ADALTokenCache(signInUserId));
            var result = await authContext.AcquireTokenSilentAsync(graphResourceUrl, clientCredential, userIdentifier);

            return result.AccessToken;
        }

        // do the call to the Graph REST API endpoint - RAW not using the Graph Client 
        private async Task<string> DoRequest(HttpMethod method, IExcelRequest excelRequest, string url)
        {
            string result = string.Empty;

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(method, url))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    if (!string.IsNullOrEmpty(sessionId))
                        request.Headers.Add("workbook-session-id", sessionId);

                    if (excelRequest != null)
                    {
                        request.Content = new StringContent(JsonConvert.SerializeObject(excelRequest), Encoding.UTF8, "application/json");
                    }

                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            result = await response.Content.ReadAsStringAsync();
                        }
                        else
                        {
                            var reason = response.ReasonPhrase;
                        }
                    }
                }
            }

            return result;
        }
    }
}