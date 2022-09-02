// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using AdaptiveCards;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.BotBuilderSamples.Models;
using TypeaheadSearch.Models;
using System.Net.Http;
using System.Net;
using AdaptiveCardsBot;

namespace Microsoft.BotBuilderSamples.Bots
{
    public class TeamsTaskModuleBot : TeamsActivityHandler
    {
        private readonly string _baseUrl;

        public TeamsTaskModuleBot(IConfiguration config)
        {
            _baseUrl = config["BaseUrl"].EndsWith("/") ? config["BaseUrl"] : config["BaseUrl"] + "/";
        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var reply = MessageFactory.Attachment(new[] { GetTaskModuleHeroCardOptions(), GetTaskModuleAdaptiveCardOptions() });
            await turnContext.SendActivityAsync(reply, cancellationToken);
        }

        protected override Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            var asJobject = JObject.FromObject(taskModuleRequest.Data);
            var value = asJobject.ToObject<CardTaskFetchValue<string>>()?.Data;

            var taskInfo = new TaskModuleTaskInfo();
            switch (value)
            {
                case TaskModuleIds.YouTube:
                    taskInfo.Url = taskInfo.FallbackUrl = _baseUrl + TaskModuleIds.YouTube;
                    SetTaskInfo(taskInfo, TaskModuleUIConstants.YouTube);
                    break;
                case TaskModuleIds.CustomForm:
                    taskInfo.Url = taskInfo.FallbackUrl = _baseUrl + TaskModuleIds.CustomForm;
                    SetTaskInfo(taskInfo, TaskModuleUIConstants.CustomForm);
                    break;
                case TaskModuleIds.AdaptiveCard:
                    taskInfo.Card = CreateAdaptiveCardAttachment();
                    SetTaskInfo(taskInfo, TaskModuleUIConstants.AdaptiveCard);
                    break;
                default:
                    break;
            }

            return Task.FromResult(taskInfo.ToTaskModuleResponse());
        }

        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleSubmitAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            var reply = MessageFactory.Text("OnTeamsTaskModuleSubmitAsync Value: " + JsonConvert.SerializeObject(taskModuleRequest));
            await turnContext.SendActivityAsync(reply, cancellationToken);

            return TaskModuleResponseFactory.CreateResponse("");
        }

        private static void SetTaskInfo(TaskModuleTaskInfo taskInfo, UISettings uIConstants)
        {
            taskInfo.Height = uIConstants.Height;
            taskInfo.Width = uIConstants.Width;
            taskInfo.Title = uIConstants.Title.ToString();
        }

        private static Attachment GetTaskModuleHeroCardOptions()
        {
            HeroCard postAnalysisCard = new HeroCard()
            {
                Buttons = new List<CardAction>
    {
        new CardAction(ActionTypes.OpenUrl, "Show me more about this topic", value: "https://teams.microsoft.com/l/entity/c076227e-04fa-441e-ab62-8cce79c83d78/index?webUrl=https://aug24thappdeva3629atab.z13.web.core.windows.net/index.html#/tab")
    }
            };
            return postAnalysisCard.ToAttachment();

            /*
            // Create a Hero Card with TaskModuleActions for each Task Module
            return new HeroCard()
            {
                Title = "Task Module Invocation from Hero Card",
                Buttons = new[] { TaskModuleUIConstants.AdaptiveCard, TaskModuleUIConstants.CustomForm, TaskModuleUIConstants.YouTube }
                            .Select(cardType => new TaskModuleAction(cardType.ButtonTitle, new CardTaskFetchValue<string>() { Data = cardType.Id }))
                            .ToList<CardAction>(),
            }.ToAttachment();*/

        }

        private static Attachment GetTaskModuleAdaptiveCardOptions()
        {
            // Create an Adaptive Card with an AdaptiveSubmitAction for each Task Module
            var card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 2))
            {
                Body = new List<AdaptiveElement>()
                    {
                        new AdaptiveTextBlock(){ Text="Task Module Invocation from Adaptive Card", Weight=AdaptiveTextWeight.Bolder, Size=AdaptiveTextSize.Large}
                    },
                Actions = new[] { TaskModuleUIConstants.AdaptiveCard, TaskModuleUIConstants.CustomForm, TaskModuleUIConstants.YouTube }
                            .Select(cardType => new AdaptiveSubmitAction() { Title = cardType.ButtonTitle, Data = new AdaptiveCardTaskFetchValue<string>() { Data = cardType.Id } })
                            .ToList<AdaptiveAction>(),
            };

            return new Attachment() { ContentType = AdaptiveCard.ContentType, Content = card };
        }

        /// <summary>
        /// Safely casts an object to an object of type <typeparamref name="T"/> .
        /// </summary>
        /// <param name="value">The object to be casted.</param>
        /// <returns>The object casted in the new type.</returns>
        private static T SafeCast<T>(object value)
        {
            var obj = value as JObject;
            if (obj == null)
            {
                throw new InvokeResponseException(HttpStatusCode.BadRequest, $"expected type '{value.GetType().Name}'");
            }

            return obj.ToObject<T>();
        }
        private static Attachment CreateAdaptiveCardAttachment()
        {
            // combine path for cross platform support
            string[] paths = { ".", "Resources", "adaptiveCard.json" };
            var adaptiveCardJson = File.ReadAllText(Path.Combine(paths));

            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = "application/vnd.microsoft.card.adaptive",
                Content = JsonConvert.DeserializeObject(adaptiveCardJson),
            };
            return adaptiveCardAttachment;
        }

        /// <summary>
        ///  Invoked when an invoke activity is received from the connector.
        /// </summary>
        /// <param name="turnContext">The turn context.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task<InvokeResponse> OnInvokeActivityAsync(ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            InvokeResponse adaptiveCardResponse;
            if (turnContext.Activity.Name == "application/search")
            {
                var searchData = JsonConvert.DeserializeObject<DynamicSearchCard>(turnContext.Activity.Value.ToString());
                var packageResult = JObject.Parse(await (new HttpClient()).GetStringAsync($"https://azuresearch-usnc.nuget.org/query?q=id:{searchData.queryText}&prerelease=true"));
                if (packageResult == null)
                {
                    var searchResponseData = new
                    {
                        type = "application/vnd.microsoft.search.searchResponse"
                    };

                    var jsonString = JsonConvert.SerializeObject(searchResponseData);
                    JObject jsonData = JObject.Parse(jsonString);

                    adaptiveCardResponse = new InvokeResponse()
                    {
                        Status = 204,
                        Body = jsonData
                    };
                }
                else
                {
                    var packages = packageResult["data"].Select(item => (item["id"].ToString(), item["description"].ToString()));
                    var packageList = packages.Select(item => { var obj = new { title = item.Item1, value = item.Item1 + " - " + item.Item2 }; return obj; }).ToList();
                    var searchResponseData = new
                    {
                        type = "application/vnd.microsoft.search.searchResponse",
                        value = new
                        {
                            results = packageList
                        }
                    };

                    var jsonString = JsonConvert.SerializeObject(searchResponseData);
                    JObject jsonData = JObject.Parse(jsonString);

                    adaptiveCardResponse = new InvokeResponse()
                    {
                        Status = 200,
                        Body = jsonData
                    };
                }

                return adaptiveCardResponse;
            }
            else
            {
                return CreateInvokeResponse(await OnTeamsTaskModuleFetchAsync(turnContext, SafeCast<TaskModuleRequest>(turnContext.Activity.Value), cancellationToken).ConfigureAwait(false));
            }

            return null;
        }
    }
}
