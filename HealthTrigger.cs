/*================================================================================================================================

  This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.  

  THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
  INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  

  We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object 
  code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software 
  product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the 
  Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims 
  or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.

 =================================================================================================================================*/


using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Newtonsoft.Json;
using RestSharp;


namespace O365HealthMonitor
{
    public static class HealthTrigger
    {

        #region classvaribles
        private static readonly string clientId = Environment.GetEnvironmentVariable("ClientId");

        private static readonly string clientSecret = Environment.GetEnvironmentVariable("ClientSecret");

        private static readonly string domain = Environment.GetEnvironmentVariable("Domain");

        private static readonly string teamsWebhook = Environment.GetEnvironmentVariable("TeamsWebhookURL");

        private static readonly string teamsWebhookPIR = Environment.GetEnvironmentVariable("TeamsWebhookPIRURL");

        // private static readonly string teamsPlannerWebhook = "";

        private static readonly string stgAccountConnection = Environment.GetEnvironmentVariable("StorageConnectionString");

        private static ILogger logger = null;

        private static readonly HttpClient httpClient = new HttpClient();

        private static long lastBatchNumber = 20200101000000;

        private static long currentBatchNumber;

        private static CurrentStatus lastKnownSvcStatus = null;

        private static CurrentStatus currentSvcStatus = null;

        private static Messages messages = null;

        private static readonly string path = Directory.GetCurrentDirectory();
        #endregion classvaribles

        #region cloudstorage
        /// <summary>
        /// Validates the connection string information in app.config and throws an exception if it looks like 
        /// the user hasn't updated this to valid values. 
        /// </summary>
        /// <param name="storageConnectionString">The storage connection string</param>
        /// <returns>CloudStorageAccount object</returns>
        private static CloudStorageAccount CreateStorageAccountFromConnectionString(string storageConnectionString)
        {
            CloudStorageAccount storageAccount;
            try
            {
                storageAccount = CloudStorageAccount.Parse(storageConnectionString);
            }
            catch (FormatException)
            {
                logger.LogError("Invalid storage account information provided. Please confirm the AccountName and AccountKey are valid in the app.config file - then restart the sample.");
                throw;
            }
            catch (ArgumentException)
            {
                logger.LogError("Invalid storage account information provided. Please confirm the AccountName and AccountKey are valid in the app.config file - then restart the sample.");
                throw;
            }
            return storageAccount;
        }

        /// <summary>
        /// Basic operations to work with Azure Files
        /// </summary>
        /// <returns>Task</returns>
        private static async void UploadtoAzureBlobStorage(string fileName, string content)
        {
            // Retrieve storage account information from connection string
            // How to create a storage connection string - http://msdn.microsoft.com/en-us/library/azure/ee758697.aspx
            CloudStorageAccount storageAccount = CreateStorageAccountFromConnectionString(stgAccountConnection);

            // Create a blob client for interacting with the blob service.
            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();

            // Create a container for organizing blobs within the storage account.
            CloudBlobContainer container = blobClient.GetContainerReference("o365healthdata");

            logger.LogInformation($"Writing {fileName} to storage");

            try
            {
                await container.CreateIfNotExistsAsync();
            }
            catch (StorageException se)
            {
                logger.LogError("Storage Exception while trying to write files." + se.Message);
                throw;
            }

            // Upload a BlockBlob to the newly created container
            CloudBlockBlob blockBlob = container.GetBlockBlobReference(fileName);
            blockBlob.Properties.ContentType = "application/json";
            await blockBlob.UploadTextAsync(content);
        }

        #endregion cloudstorage

        #region o365messages
        /// <summary>
        /// gets token using client id and secret
        /// </summary>
        /// <returns></returns>
        static async Task<HttpResponseMessage> GetTokenAsync()
        {

            var formVars = new Dictionary<string, string>();
            formVars.Add("grant_type", "client_credentials");
            formVars.Add("resource", "https://manage.office.com");
            formVars.Add("client_id", clientId);
            formVars.Add("client_secret", clientSecret);

            var content = new FormUrlEncodedContent(formVars);

            HttpResponseMessage response = await httpClient.PostAsync("https://login.microsoftonline.com/" + domain + "/oauth2/token?api-version=1.0", content);
            response.EnsureSuccessStatusCode();
            logger.LogInformation($"Got Token {clientId}");
            return response;
        }

        /// <summary>
        /// Gets the service message from https://manage.office.com/api/v1.0/ServiceComms/Messages
        /// </summary>
        /// <param name="tok">Access Token</param>
        /// <returns></returns>
        static async Task<HttpResponseMessage> GetServiceMessage(Token tok)
        {
            logger.LogInformation("Inside getServiceMessage");
            HttpResponseMessage response = null;
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue(tok.token_type, tok.access_token);
            response = await httpClient.GetAsync("https://manage.office.com/api/v1.0/" + domain + "/ServiceComms/Messages");
            response.EnsureSuccessStatusCode();
            logger.LogInformation("got service message");

            string svcResponse = response.Content.ReadAsStringAsync().Result;

            messages = JsonConvert.DeserializeObject<Messages>(svcResponse);
                       
            logger.LogInformation($"Service Message Count {messages.value.Count()}");
            return response;
        }

        /// <summary>
        /// gets the current status from https://manage.office.com/api/v1.0/ServiceComms/CurrentStatus
        /// </summary>
        /// <param name="tok">Access Token</param>
        /// <returns></returns>
        static async Task<HttpResponseMessage> getCurrentStatus(Token tok)
        {
            logger.LogInformation("Inside getCurrentStatus");
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue(tok.token_type, tok.access_token);

            HttpResponseMessage response = await httpClient.GetAsync("https://manage.office.com/api/v1.0/" + domain + "/ServiceComms/CurrentStatus");
            response.EnsureSuccessStatusCode();
            return response;
        }

        #endregion o365messages


        /// <summary>
        /// Azure Functions uses the NCronTab library to interpret NCRONTAB expressions. 
        /// Currently, set to run every minute.
        /// An NCRONTAB expression is similar to a CRON expression except that it includes 
        /// an additional sixth field at the beginning to use for time precision in seconds: 
        /// {second} {minute} {hour} {day} {month} {day-of-week}
        /// https://docs.microsoft.com/en-us/azure/azure-functions/functions-bindings-timer?tabs=csharp#ncrontab-expressions
        /// <param name="myTimer"></param>
        /// <param name="log"></param>
        [FunctionName("HealthTrigger")]
        public static void Run([TimerTrigger("%TimerSchedule%")]TimerInfo myTimer, ILogger log)
        {

            logger = log;
            currentBatchNumber = long.Parse(DateTime.UtcNow.ToString("yyyyMMddHHmmss"));
            if (httpClient.BaseAddress == null)
            {
                httpClient.BaseAddress = new Uri("https://localhost/");
            }

            logger.LogInformation("Inside Run");
            if (Environment.GetEnvironmentVariable("Env") == "Dev")
            {
                long.TryParse(System.IO.File.ReadAllText(path + @"\\Data\\LastRun.json"), out lastBatchNumber);
            }
            else
            {
                long.TryParse(System.IO.File.ReadAllText(@"D:\\home\\site\\wwwroot\\Data\\LastRun.json"), out lastBatchNumber);
            }

            logger.LogInformation($"Got LastBatchNumber {lastBatchNumber.ToString()}, entering Init..");

            Init();

            RunAsync().GetAwaiter().GetResult();

            // Processing completed, dump the data into json
            if (Environment.GetEnvironmentVariable("Env") == "Dev")
            {
                System.IO.File.WriteAllText(path + @"\\Data\\LastRun.json", currentBatchNumber.ToString());
            }
            else
            {
                System.IO.File.WriteAllText(@"D:\\home\\site\\wwwroot\\Data\\LastRun.json", currentBatchNumber.ToString());
            }

            log.LogInformation($"C# {lastBatchNumber} Timer trigger function executed at: {DateTime.Now}");
        }

        static void Init()
        {
            try
            {
                logger.LogInformation("Inside Init");
                // Get the Message Corresponding to last batch from file system
                if (Environment.GetEnvironmentVariable("Env") == "Dev")
                {
                    lastKnownSvcStatus = JsonConvert.DeserializeObject<CurrentStatus>(System.IO.File.ReadAllText(path + $"\\Data\\CurrentStatus.json"));
                }
                else
                {
                    lastKnownSvcStatus = JsonConvert.DeserializeObject<CurrentStatus>(System.IO.File.ReadAllText($"D:\\home\\site\\wwwroot\\Data\\CurrentStatus.json"));
                }

                logger.LogInformation($"lastKnownSvcStatus {lastKnownSvcStatus}");

            }
            catch (Exception ex)
            {
                logger.LogInformation($"Could not load lastKnownSvcStatus {ex.Message}");
                logger.LogInformation($"lastKnownSvcStatus {ex.StackTrace.ToString()}");
            }
        }

        static async Task RunAsync()
        {

            logger.LogInformation("Inside RunAsync");
            // Get Token
            var httpResponse = await GetTokenAsync();
            string apiResponse = httpResponse.Content.ReadAsStringAsync().Result;
            var token = JsonConvert.DeserializeObject<Token>(apiResponse);

            //Get Service Message
            var currentStatus = await getCurrentStatus(token);
            string currResponse = currentStatus.Content.ReadAsStringAsync().Result;
            currentSvcStatus = JsonConvert.DeserializeObject<CurrentStatus>(currResponse);
            logger.LogInformation($"Status Message Count {currentSvcStatus.value.Count()}");
            await GetServiceMessage(token);

            NotifyPIR();

            //Compare with last known status and notify if needed
            CompareAndNotify();


            
            // notifyTeamsPlanner(messages);

            if (Environment.GetEnvironmentVariable("Env") == "Dev")
            {
                System.IO.File.WriteAllText(path + $"\\Data\\CurrentStatus.json", JsonConvert.SerializeObject(currentSvcStatus));
                System.IO.File.WriteAllText(path + $"\\Data\\Messages.json", JsonConvert.SerializeObject(messages));
            }
            else
            {
                UploadtoAzureBlobStorage("CurrentStatus.json", JsonConvert.SerializeObject(currentSvcStatus));
                UploadtoAzureBlobStorage("Messages.json", JsonConvert.SerializeObject(messages));
                UploadtoAzureBlobStorage("LastRun.json", currentBatchNumber.ToString());

                System.IO.File.WriteAllText($"D:\\home\\site\\wwwroot\\Data\\CurrentStatus.json", JsonConvert.SerializeObject(currentSvcStatus));
                System.IO.File.WriteAllText($"D:\\home\\site\\wwwroot\\Data\\Messages.json", JsonConvert.SerializeObject(messages));


            }
        }

        private static void NotifyPIR()
        {
            CultureInfo provider = CultureInfo.InvariantCulture;
            // It throws Argument null exception  
            DateTime dateTime = DateTime.ParseExact(lastBatchNumber.ToString(), "yyyyMMddHHmmss", provider);

            List<MessageValue> lst = messages.value.Where(a => (a.Status == "Post-incident report published" &&  a.LastUpdatedTime.ToUniversalTime() > dateTime)).ToList();

            foreach (var item in lst)
            {
                notifyTeamsPIR(item);
            }
        }

        static void notifyTeamsPIR(MessageValue message)
        {

            // My Teams LInk
            var client = new RestClient(teamsWebhookPIR);
            var request = new RestRequest(Method.POST);
            request.AddHeader("cache-control", "no-cache");
            request.AddHeader("Connection", "keep-alive");
            request.AddHeader("accept-encoding", "gzip, deflate");
            request.AddHeader("Host", "outlook.office.com");
            request.AddHeader("Cache-Control", "no-cache");
            request.AddHeader("Accept", "*/*");
            request.AddHeader("Content-Type", "application/json");

            List<Fact> facts = new List<Fact>();
            facts.Add(new Fact() { Name = "Title", Value = message.Title });
            facts.Add(new Fact() { Name = "Id", Value = message.Id });
            facts.Add(new Fact() { Name = "Severity", Value = message.Severity });
            facts.Add(new Fact() { Name = "Start Time", Value = message.StartTime.ToString() });

            if (message.EndTime != null)
            {
                facts.Add(new Fact() { Name = "End Time", Value = message.EndTime.ToString() });
            }

            foreach (var item in message.Messages)
            {
                facts.Add(new Fact() { Name = item.PublishedTime.ToString(), Value = item.MessageText });
            }

            Section section = new Section() {
                ActivityTitle = message.Id + " - " + message.Title,
                ActivitySubtitle = message.WorkloadDisplayName,
                ActivityImage = "https://teamsnodesample.azurewebsites.net/static/img/image5.png",
                Markdown = true,
                Facts = facts
            };

            List<Target> targets = new List<Target>();
            targets.Add(new Target() { Os = "default", Uri = message.PostIncidentDocumentUrl });

            //List<Section> sections = new List<Section>();
            //sections.Add(section);

            PotentialAction potentialAction = new PotentialAction() { 
             Type = "OpenUri",
             Name = "Post Incident Document Url",
             Targets = targets
            };

            List<PotentialAction> potentialActions = new List<PotentialAction>() { potentialAction };
              Root root = new Root(){ 
               Type = "MessageCard",
               Context = "http://schema.org/extensions",
               ThemeColor = "0076D7",
               Summary = message.Status,
               Sections = new List<Section>() { section},
               PotentialAction = new List<PotentialAction>() { potentialAction }
            };

            request.AddParameter("undefined", JsonConvert.SerializeObject(root) , ParameterType.RequestBody);
            IRestResponse response = client.Execute(request);
            //if (response.StatusCode != StatusCodes.Status200OK)
            //{
                logger.LogInformation(response.ErrorMessage);
             if (response.ErrorException != null) { 
                logger.LogInformation(response.ErrorException.ToString());
            }
        }

        /// <summary>
        /// Compares previous degradations and picksup services degraded
        /// </summary>
        static void CompareAndNotify()
        {

            List<FeatureValue> serviceDegraded = new List<FeatureValue>();

            foreach (FeatureValue item in currentSvcStatus.value)
            {
                if (item.FeatureStatus.Where(a => a.FeatureServiceStatus == "ServiceDegradation").Count() > 0)
                {
                    serviceDegraded.Add(item);
                }
            }

            List<FeatureValue> serviceDegradedNotified = new List<FeatureValue>();

            foreach (FeatureValue item in lastKnownSvcStatus.value)
            {
                if (item.FeatureStatus.Where(a => a.FeatureServiceStatus == "ServiceDegradation").Count() > 0)
                    serviceDegradedNotified.Add(item);
            }

            List<FeatureStatus> dStatusLast = new List<FeatureStatus>();
            List<FeatureStatus> dStatusNew = new List<FeatureStatus>();

            foreach (var item in serviceDegradedNotified)
            {
                dStatusLast.AddRange(item.FeatureStatus);
            }

            foreach (var item in serviceDegraded)
            {
                dStatusNew.AddRange(item.FeatureStatus);
            }

            var degrades = dStatusNew.Select(x => x.FeatureName).ToArray().Except(dStatusLast.Select(x => x.FeatureName).ToArray());

            if (degrades.Count() > 0)
            {
                foreach (var itemcol in serviceDegraded)
                {
                    StringBuilder builder = new StringBuilder("<table style='width:100%;'>");
                    builder.Append("<tr style='background-color:#c0c0c0;'><td style='border-style:solid;border-width:1px'><b>" + itemcol.WorkloadDisplayName + "</b></td><td style='border-style:solid;border-width:1px'><b>" + itemcol.StatusDisplayName + "</b></td></tr>");
                    foreach (FeatureStatus item in itemcol.FeatureStatus)
                    {
                        builder.Append("<tr><td style='border-style:solid;border-width:1px'>" + item.FeatureDisplayName + "</td><td style='border-style:solid;border-width:1px'>" + item.FeatureServiceStatusDisplayName + "</td></tr>");
                    }
                    builder.Append("</table>");

                    notifyTeams(builder.ToString());
                }
            }
        }

        /// <summary>
        /// Method notifies the teams channel provided, when new service degradations occur
        /// </summary>
        /// <param name="message"></param>
        static void notifyTeams(string message)
        {
            // My Teams LInk
            var client = new RestClient(teamsWebhook);
            var request = new RestRequest(Method.POST);
            request.AddHeader("cache-control", "no-cache");
            request.AddHeader("Connection", "keep-alive");
            request.AddHeader("accept-encoding", "gzip, deflate");
            request.AddHeader("Host", "outlook.office.com");
            request.AddHeader("Cache-Control", "no-cache");
            request.AddHeader("Accept", "*/*");
            request.AddHeader("Content-Type", "application/json");
            request.AddParameter("undefined", "{\r\n  \"@context\": \"https://schema.org/extensions\",\r\n  \"@type\": \"MessageCard\",\r\n  \"themeColor\": \"0072C6\",\r\n  \"title\": \"Service Degradation\",\r\n  \"text\": \"Click **Learn More** to learn more ! <br/>" + message + "\",\r\n  \"potentialAction\": [\r\n    {\r\n      \"@type\": \"OpenUri\",\r\n      \"name\": \"Learn More\",\r\n      \"targets\": [\r\n        { \"os\": \"default\", \"uri\": \"https://admin.microsoft.com/Adminportal/Home?ref=MessageCenter\" }\r\n      ]\r\n    }\r\n  ]\r\n}", ParameterType.RequestBody);
            IRestResponse response = client.Execute(request);
            //if (response.StatusCode != StatusCodes.Status200OK)
            //{
            if (response.ErrorException != null)
            {
                logger.LogInformation(response.ErrorMessage);
                logger.LogInformation(response.ErrorException.ToString());
            }
        }


        #region NOT USED
        static void notifyTeamsPlanner(Messages value)
        {
            List<MessageValue> val = value.value;

            //List<CardSection> cardSection = new List<CardSection>();
            MessageCard card = new MessageCard();


            foreach (var item in val)
            {
                List<Message> msg = new List<Message>();
                msg = item.Messages;

                List<CardFacts> facts = new List<CardFacts>();

                foreach (var ite in msg)
                {
                    CardFacts fa = new CardFacts()
                    {
                        name = ite.PublishedTime.ToString(),
                        value = ite.MessageText
                    };
                    facts.Add(fa);
                }
                List<CardSection> cardSection = new List<CardSection>();

                CardSection cardS = new CardSection()
                {
                    facts = facts,
                    text = item.WorkloadDisplayName,
                    activityTitle = item.FeatureDisplayName,
                    activitySubtitle = item.Id + " " + item.ImpactDescription
                };

                cardSection.Add(cardS);

                card = new MessageCard()
                {
                    sections = cardSection,
                    summary = item.Status,
                    title = item.Title
                };

                // My Teams LInk
                var client = new RestClient(teamsWebhook);
                var request = new RestRequest(Method.POST);
                request.AddHeader("cache-control", "no-cache");
                request.AddHeader("Connection", "keep-alive");
                request.AddHeader("accept-encoding", "gzip, deflate");
                request.AddHeader("Host", "outlook.office.com");
                request.AddHeader("Cache-Control", "no-cache");
                request.AddHeader("Accept", "*/*");
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter(JsonConvert.SerializeObject(card), ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
            }
        }

        #endregion

    }
}
