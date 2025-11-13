using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace ImportExcel_With_ASPNETMVC.Controllers
{
    public class ChatGPTIntegrationController : Controller
    {
        private static readonly string apiUrl = "https://api.openai.com/v1/chat/completions";
        // GET: ChatGPTIntegration
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<ActionResult> GetChatGPTResponse(string prompt)
        {
            var response = await GetChatGPTResponseAsync(prompt);
            return Json(response);
        }

        private async Task<string> GetChatGPTResponseAsync(string prompt)
        {
            var apiKey = ConfigurationManager.AppSettings["ChatgptapiKey"].ToString();

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);

                var requestBody = new
                {
                    model = "gpt-3.5-turbo",
                    messages = new[]
                    {
                        new{role="system",content="you are helpful assistent"},
                        new{role="user",content=prompt}
                    },
                    max_tokens = 100
                };

                var content = new StringContent(JsonConvert.SerializeObject(requestBody), System.Text.Encoding.UTF8, "application/json");

                var response = await client.PostAsync(apiUrl, content);
                response.EnsureSuccessStatusCode();

                var responseBody = await response.Content.ReadAsStringAsync();
                dynamic result = JsonConvert.DeserializeObject(responseBody);

                return result.choices[0].message.content.toString();
            }
        }
    }
}