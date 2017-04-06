using Microsoft.Bot.Connector.DirectLine;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace MSBotTest
{
    /// <summary>
    /// channel data class to pass client activity id in ChannelData parameter
    /// </summary>
    public class ChannelData
    {
        public string clientActivityId { get; set; }
    }

    class Program
    {
        private static string directLine = GetConfig("Channel");
        private static string messageType = GetConfig("MessageType");
        private static string englishLocale = GetConfig("Locale");
        private static string botName = GetConfig("BotName");
        private static string recordedText = string.Empty;
        private static string botId = GetConfig("BotId");
        private static string serviceUrl = GetConfig("ServiceUrl");
        private static string directLineSecret = GetConfig("DirectLineSecret");


        /// <summary>
        /// Main executable method
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            // get the test steps
            string testCaseFile = GetConfig("TestCaseFile");
            TestSuite testSuite = new TestSuite(testCaseFile);

            int passCount = 0;
            int failCount = 0;

            foreach (var testCase in testSuite.TestCases)
            {
                // execute the test suite
                foreach (var testStep in testCase.Steps)
                {
                    // to process the test cases in parallel, concepts like synchronization needs to be introduced
                    // this is a basic code on how to write functional regression test scenario
                    ProcessTestCase(ref passCount, ref failCount, testCase, testStep);
                }
            }
            Console.WriteLine("");

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Passed: {passCount}");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Failed: {failCount}");

            Console.ReadLine();
        }

        /// <summary>
        /// Process the test case
        /// </summary>
        /// <param name="passCount"></param>
        /// <param name="failCount"></param>
        /// <param name="testCase"></param>
        /// <param name="testStep"></param>
        private static void ProcessTestCase(ref int passCount, ref int failCount, TestCase testCase, TestCaseStep testStep)
        {
            bool testCaseResult = true;
            Console.WriteLine($"Starting Test case: {testCase.SheetName} {testStep.Number}");

            Microsoft.Bot.Connector.DirectLine.DirectLineClient client = new Microsoft.Bot.Connector.DirectLine.DirectLineClient(directLineSecret);
            var conversation = client.Conversations.StartConversationWithHttpMessagesAsync().GetAwaiter().GetResult();
            string conversationId = conversation.Body.ConversationId;
            string clientActivtyId = string.Format("{0}.{1}", GenerateNumber(13), GenerateNumber(17));
            int watermark = 0;
            string from = "tester";

            int i = 0;
            foreach (var action in testStep.Actions)
            {
                var messagePostResponse = client.Conversations.PostActivityWithHttpMessagesAsync(conversationId, new Microsoft.Bot.Connector.DirectLine.Activity(
                   type: messageType,
                   id: GenerateNumber(17),
                   text: action.Input, // replace with text recieved from voice to text done above via tts service
                   channelId: directLine,
                   timestamp: DateTime.Now,
                   localTimestamp: DateTime.Now,
                   serviceUrl: serviceUrl,
                   fromProperty: new ChannelAccount(from, from),
                   conversation: new ConversationAccount() { Id = conversationId },
                   recipient: new ChannelAccount() { Id = botId, Name = botName },
                   locale: englishLocale,
                   channelData: JsonConvert.SerializeObject(new ChannelData() { clientActivityId = GetClientActivityToResponse(clientActivtyId, watermark) })
                   )).Result;

                if (i == 0) { i++; }
                else
                {
                    watermark++;
                }


                DateTime dt = DateTime.Now;
                bool resultFound = false;

                var responseText = new List<string>();
                // check for 15 seconds max and if no response recieved raise exception
                while (true && dt.AddSeconds(15) >= DateTime.Now)
                {
                    var activities = client.Conversations.GetActivitiesWithHttpMessagesAsync(conversationId, watermark.ToString()).Result;
                    foreach (var activity in activities.Body.Activities)
                    {
                        //if (activity.From.Id == botId)
                        if (activity.Attachments.Count > 0)
                        {
                            switch (activity.Attachments[0].ContentType)
                            {
                                case "application/vnd.microsoft.card.hero":
                                    var heroCard = JsonConvert.DeserializeObject<HeroCard>(Convert.ToString(activity.Attachments[0].Content));
                                    string heroCardText = heroCard.Text.Trim();
                                    heroCardText += "(";
                                    foreach (var button in heroCard.Buttons)
                                    {
                                        heroCardText += $"{button.Title}|";
                                    }
                                    heroCardText = heroCardText.Substring(0, heroCardText.Length - 1);
                                    heroCardText += ")";
                                    responseText.Add(heroCardText);
                                    break;
                                    // add more cards
                            }
                        }
                        else
                        {
                            responseText.Add(activity.Text);
                        }

                        watermark++;
                        //check against hero card, login card etc...activity.Attachments[0].ContentType
                        resultFound = true;
                    }
                    if (resultFound)
                    {
                        break;
                    }
                    System.Threading.Thread.Sleep(1000); // change the time out accordingly
                }
                if (resultFound)
                {
                    string expectedResponse = action.ExpectedResponse;

                    Regex regexCompleteInput = new Regex(@"\$\{(\d)+\}");
                    var m = regexCompleteInput.Matches(expectedResponse);
                    for (int c = 0; c < m.Count; c++)
                    {
                        var entityDefinition = m[c].Value;
                        int inputNumber = Convert.ToInt32(entityDefinition.Substring(2, entityDefinition.Length - 3));

                        expectedResponse = expectedResponse.Replace(entityDefinition,
                                testStep.Actions.Where(a => a.StepNumber == inputNumber).First().Input);
                    }

                    Regex regexEntityInput = new Regex(@"\$\{(\d)+\-(\d)+\}");
                    m = regexEntityInput.Matches(expectedResponse);
                    for (int c = 0; c < m.Count; c++)
                    {
                        var entityDefinition = m[c].Value;
                        string expression = entityDefinition.Substring(2, entityDefinition.Length - 3);
                        int inputNumber = Convert.ToInt32(expression.Split(new char[] { '-' })[0]);
                        int entityIndex = Convert.ToInt32(expression.Split(new char[] { '-' })[1]);

                        expectedResponse = expectedResponse.Replace(entityDefinition,
                                testStep.Actions.Where(a => a.StepNumber == inputNumber).First()
                                .Entities.Where(e => e.Index == entityIndex).First().Name);
                    }

                    if (string.Join(Environment.NewLine, responseText.ToArray()).ToLower().Replace("\r", string.Empty).Replace("\n", string.Empty)
                        == action.ExpectedResponse.ToLower().Replace("\r", string.Empty).Replace("\n", string.Empty))
                    {
                        // success
                        //Console.WriteLine("success");
                        // do nothing
                    }
                    else
                    {
                        // failure
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Failure: {testStep.Number}, step: {action.StepNumber}");
                        Console.WriteLine($"Expected output: {action.ExpectedResponse.ToLower()}");
                        Console.WriteLine($"Actual output: {string.Join(Environment.NewLine, responseText.ToArray()).ToLower().Replace("\r\n", "\n")}");
                        Console.ForegroundColor = ConsoleColor.White;
                        testCaseResult = false;
                        break;
                    }
                }
            }
            if (testCaseResult)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Test case passed: {testCase.SheetName} {testStep.Number}");
                Console.ForegroundColor = ConsoleColor.White;
                passCount++;
            }
            else
            {
                failCount++;
            }
            Console.WriteLine("*******");
        }

        /// <summary>
        /// Generate the activity id along with incremental watermark
        /// </summary>
        /// <param name="clientActivityId">client activity id</param>
        /// <param name="watermark">watermark</param>
        /// <returns></returns>
        private static string GetClientActivityToResponse(string clientActivityId, int watermark)
        {
            return string.Format("{0}.{1}", clientActivityId, watermark);
        }

        /// <summary>
        /// Generate a random digit of <paramref name="n"/> charaters
        /// </summary>
        /// <param name="n">lenght of digit</param>
        /// <returns></returns>
        private static string GenerateNumber(int n)
        {
            Random random = new Random();
            string r = "";
            int i;
            for (i = 1; i < n; i++)
            {
                r += random.Next(0, 9).ToString();
            }
            return r;
        }

        /// <summary>
        /// get the value from configuration
        /// </summary>
        /// <param name="name">configuration key name</param>
        /// <returns></returns>
        private static string GetConfig(string name)
        {
            return System.Configuration.ConfigurationManager.AppSettings[name] as string;
        }

    }
}
