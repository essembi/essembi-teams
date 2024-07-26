using AdaptiveCards;
using Essembi.Integrations.Teams.Model;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using Microsoft.Extensions.Configuration;
using RestSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace Essembi.Integrations.Teams.Bots
{
    public class TeamsMessagingExtensionsActionBot : TeamsActivityHandler
    {
        readonly string _teamsKey;
        readonly UserState _userState;

        public TeamsMessagingExtensionsActionBot(IConfiguration configuration, UserState userState) 
            : base()
        {

            _teamsKey = configuration["TeamsIntegrationKey"];
            _userState = userState;
        }

        #region Message / Compose Actions

        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionSubmitActionAsync(
            ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            if (action.CommandId != "createTicket" && action.CommandId != "createTicketMessage")
            {
                return new MessagingExtensionActionResponse();
            }

            //-- Data should be a JObject.
            if (action.Data is not Newtonsoft.Json.Linq.JObject data)
            {
                return ErrorMessage("An unexpected error occurred while processing the input. Please try again later. Contact Essembi support if this problem persists.");
            }

            if (data.ContainsKey("environment"))
            {
                //-- User chose environment. Show the ticket entry window.
                //-- Get the auth data from the user state.
                await _userState.LoadAsync(turnContext, cancellationToken: cancellationToken);
                var authData = _userState.Get(turnContext)[nameof(EssembiAuthResponse)];
                var deserialized = authData.ToObject<EssembiAuthResponse>();
                if (authData == null || deserialized == null)
                {
                    return ErrorMessage("Your session has expired. Please try again.");
                }

                //-- Select the correct environment.
                var appId = long.Parse(data["environment"].ToString());
                var app = deserialized.Apps.FirstOrDefault(v => v.Id == appId);
                if (app == null)
                {
                    return ErrorMessage("An unexpected error occurred while processing the input. Please try again later. Contact Essembi support if this problem persists.");
                }

                //-- Get the message body.
                var email = _userState.Get(turnContext)["email"]?.ToString();
                var subject = _userState.Get(turnContext)["subject"]?.ToString();
                var body = _userState.Get(turnContext)["body"]?.ToString();

                //-- Show the window.
                return ShowTicketEntryWindow(email, subject, body, app);
            }
            else if (!data.ContainsKey("appId"))
            {
                //-- Probably JIT installation.
                return new MessagingExtensionActionResponse();
            }

            //-- User submitted the ticket.
            //-- Set up the request.
            var requestBody = new CreateFromMSTeamsRequest();
            requestBody.Email = data["email"].ToString();
            requestBody.AppId = long.Parse(data["appId"].ToString());
            requestBody.TableId = long.Parse(data["tableId"].ToString());
            requestBody.Values = new Dictionary<string, object>();

            foreach (var prop in data.Properties())
            {
                if (prop.Name == "appId" || prop.Name == "tableId")
                {
                    continue;
                }

                //-- Turn newlines into BRs for multi-line inputs.
                var stringVal = prop.Value.ToString();
                if (stringVal.Contains('\r') || stringVal.Contains('\n'))
                {
                    var split = stringVal.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.TrimEntries);
                    stringVal = string.Join("<br />", split);
                }

                requestBody.Values[prop.Name] = stringVal;
            }

            using var client = MakeRequest("Integrations/MSTeams/Create", out var request);

            //-- Add the request body.
            request.AddJsonBody(requestBody);

            //-- Execute the request.
            var response = await client.ExecutePostAsync(request, cancellationToken);
            if (response.StatusCode != System.Net.HttpStatusCode.OK)
            {
                //-- Report the problem to the user.
                string message = null;
                if (response.Content != null)
                {
                    message = JsonSerializer.Deserialize<MessageResponse>(response.Content)?.Message;
                }

                if (string.IsNullOrEmpty(message))
                {
                    message = "An error occurred while trying to create the ticket. Please try again later. Contact Essembi support if this problem persists.";
                }

                return ErrorMessage(message);
            }

            //-- Return the success message.
            var successMessage = JsonSerializer.Deserialize<CreateFromMSTeamsResult>(response.Content!);
            var heroCard = new HeroCard
            {
                Subtitle = $"Summary: {successMessage.Name}",
                Text = "A ticket has been created successfully. You may now view this ticket in Essembi.",
                Tap = new CardAction
                {
                    Type = ActionTypes.OpenUrl,
                    Title = "View Ticket in Essembi",
                    Value = successMessage.Url
                },
                Buttons = new List<CardAction>
                {
                    new CardAction
                    {
                        Type = ActionTypes.OpenUrl,
                        Title = "View Ticket in Essembi",
                        Value = successMessage.Url
                    }
                }
            };

            if (string.IsNullOrEmpty(successMessage.Number))
            {
                heroCard.Title = $"Ticket has been created!";
            }
            else
            {
                heroCard.Title = $"Ticket #{successMessage.Number} has been created!";
            }

            return new MessagingExtensionActionResponse
            {
                ComposeExtension = new MessagingExtensionResult
                {
                    Type = "result",
                    AttachmentLayout = "list",
                    Attachments = new List<MessagingExtensionAttachment>()
                        {
                            new MessagingExtensionAttachment
                            {
                                Content = heroCard,
                                ContentType = HeroCard.ContentType,
                                Preview = heroCard.ToAttachment(),
                            },
                        },
                },
            };
        }

        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionFetchTaskAsync(
            ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            if (action.CommandId != "createTicket" && action.CommandId != "createTicketMessage")
            {
                return new MessagingExtensionActionResponse();
            }

            string subject = null, body = null;
            if (action.MessagePayload != null)
            {
                subject = action.MessagePayload.Subject;
                body = action.MessagePayload.Body?.Content;

                if (body != null)
                {
                    //-- Remove the HTML tags.
                    body = RemoveMarkup(body);
                }
            }

            return await CreateTicket(turnContext, cancellationToken, subject, body);
        }

        static string RemoveMarkup(string text)
        {
            //-- Remove the HTML tags.
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(text);
            text = doc.DocumentNode.InnerText;
            return text;
        }

        private async Task<MessagingExtensionActionResponse> CreateTicket(ITurnContext turnContext, CancellationToken cancellationToken, string subject = null, string body = null)
        {
            var email = await GetEmail(turnContext, cancellationToken);
            if (email.response != null)
            {
                return email.response;
            }

            var member = email.account;

            using var client = MakeRequest("Integrations/MSTeams/Authenticate", out var request);

            //-- Add the teams-authenticated email to match to an essembi user.
            request.AddBody(new
            {
                email = member.Email
            });

            //-- Execute the request.
            var response = await client.ExecutePostAsync(request, cancellationToken);
            switch (response.StatusCode)
            {
                case System.Net.HttpStatusCode.OK:
                    break;

                case System.Net.HttpStatusCode.NotFound:
                    return ErrorMessage("You do not have an Essembi account. Sign up today at essembi.com!");

                default:
                    return ErrorMessage("An error occurred while trying to authenticate your account. Please try again later. Contact Essembi support if this problem persists.");
            }

            //-- Deserialize with the defined message contract.
            var deserialized = JsonSerializer.Deserialize<EssembiAuthResponse>(response.Content!);
            if (deserialized == null)
            {
                return ErrorMessage("An unexpected error occurred. Please try again later. Contact Essembi support if this problem persists.");
            }

            if (deserialized.Apps == null || deserialized.Apps.Length == 0)
            {
                //-- No apps are enabled.
                return ErrorMessage("You must enable the Teams integration in Essembi. This is done in Settings > Integrations.");
            }
            else if (deserialized.Apps.Length == 1)
            {
                //-- Only one app is enabled. Show the entry window.
                var app = deserialized.Apps[0];
                return ShowTicketEntryWindow(member.Email, subject, body, app);
            }

            //-- Multiple apps are enabled. Need to show the app selection window.
            //-- Store the needed data in memory state for now.
            await _userState.ClearStateAsync(turnContext, cancellationToken);

            var authProp = _userState.CreateProperty<EssembiAuthResponse>(nameof(EssembiAuthResponse));
            await authProp.SetAsync(turnContext, deserialized, cancellationToken);

            var subjectProp = _userState.CreateProperty<string>("subject");
            await subjectProp.SetAsync(turnContext, subject, cancellationToken);

            var bodyProp = _userState.CreateProperty<string>("body");
            await bodyProp.SetAsync(turnContext, body, cancellationToken);

            await _userState.SaveChangesAsync(turnContext, cancellationToken: cancellationToken);

            var card = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>()
                {
                    new AdaptiveTextBlock
                    {
                        Text = "You have access to multiple Essembi environments. Select the environment you want to create a ticket in.",
                        Wrap = true
                    },
                    new AdaptiveChoiceSetInput
                    {
                        Label = "Select Environment",
                        IsRequired = true,
                        Id = "environment",
                        Choices = deserialized.Apps.Select(v => new AdaptiveChoice
                        {
                            Title = v.Name,
                            Value = v.Id.ToString()
                        }).ToList()
                    }
                },
                Actions = new List<AdaptiveAction>()
                {
                    new AdaptiveSubmitAction
                    {
                        Title = "Submit",
                        Type = "Action.Submit",
                        Data = new Dictionary<string, object>
                        {
                            { "email", member.Email }
                        }
                    }
                }
            };

            //-- Return the form.
            return new MessagingExtensionActionResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo
                    {
                        Card = new Attachment
                        {
                            ContentType = "application/vnd.microsoft.card.adaptive",
                            Content = card
                        },
                        Height = "small",
                        Width = "small",
                        Title = "Choose Your Essembi Environment",
                    },
                },
            };
        }

        private static async Task<(MessagingExtensionActionResponse response, TeamsChannelAccount account)> GetEmail(ITurnContext turnContext, CancellationToken cancellationToken)
        {
            try
            {
                //-- Get member information.
                var member = await TeamsInfo.GetMemberAsync(turnContext, turnContext.Activity.From.Id, cancellationToken);

                return (null, member);
            }
            catch (ErrorResponseException ex)
            {
                if (ex.Body.Error.Code == "BotNotInConversationRoster" || 
                    ex.Response.StatusCode == System.Net.HttpStatusCode.Forbidden)
                {
                    //-- The bot is not in the conversation roster.
                    var resp = new MessagingExtensionActionResponse
                    {
                        Task = new TaskModuleContinueResponse
                        {
                            Value = new TaskModuleTaskInfo
                            {
                                Card = GetAdaptiveCardAttachmentFromFile("justintimeinstallation.json"),
                                Height = 200,
                                Width = 400,
                                Title = "App Installation",
                            },
                        },
                    };

                    return (resp, null);
                }
                else if (ex.Response.StatusCode == System.Net.HttpStatusCode.BadRequest)
                {
                    //-- There is a bug in teams where you can not use GetMemberAsync until the bot has been added to a chat or channel.
                    //-- This means that ticket submission can not be used in a 1:1 chat (with the bot specifically) until the bot has been added to a chat or channel.
                    var resp = new MessagingExtensionActionResponse
                    {
                        Task = new TaskModuleContinueResponse
                        {
                            Value = new TaskModuleTaskInfo
                            {
                                Card = GetAdaptiveCardAttachmentFromFile("botnotready.json"),
                                Height = 200,
                                Width = 400,
                                Title = "Add Bot to Chat or Channel",
                            },
                        },
                    };

                    return (resp, null);
                }

                //-- Log some additional information to go along with the exception in the error log.
                Console.WriteLine($"{ex.GetType()}: {ex.Message} -- {ex.Response.Content}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"{ex.InnerException.GetType()}: {ex.InnerException.Message}");
                }

                var sb = new StringBuilder();
                sb.Append("Role: ").AppendLine(turnContext.Activity.From.Role);
                sb.Append("Name: ").AppendLine(turnContext.Activity.From.Name);
                sb.Append("Id: ").AppendLine(turnContext.Activity.From.Id);
                if (turnContext.Activity.From.Properties != null)
                {
                    foreach (var prop in turnContext.Activity.From.Properties)
                    {
                        sb.Append("Properties.").Append(prop.Key).Append(": ").AppendLine(prop.Value?.ToString());
                    }
                }

                Console.WriteLine(sb.ToString());

                //-- It's a different error.
                throw;
            }
        }

        #endregion

        #region Scope Message Actions

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            //-- Get the message text.
            var text = RemoveMarkup(turnContext.Activity.Text?.Trim() ?? string.Empty);
            var lowerText = text.ToLowerInvariant();

            //-- Trim off the @-message in group texts and channels.
            if (lowerText.StartsWith('@'))
            {
                text = text.Substring(1).Trim();
                lowerText = lowerText.Substring(1).Trim();
            }

            if (lowerText.StartsWith("essembi", StringComparison.InvariantCulture))
            {
                text = text.Substring(7).Trim();
                lowerText = lowerText.Substring(7).Trim();
            }

            if (lowerText.StartsWith("search", StringComparison.InvariantCulture))
            {
                var email = await GetEmail(turnContext, cancellationToken);
                if (email.response != null)
                {
                    //-- This should never happen. Return error message.
                    await turnContext.SendActivityAsync(
                        MessageFactory.Text($"The search failed because the bot is not present in this channel or conversation."), cancellationToken);
                    return;
                }

                //-- Sanitize the search string.
                var query = text.Substring(6).Trim();
                if(string.IsNullOrEmpty(query))
                {
                    //-- Return an error message.
                    await turnContext.SendActivityAsync(
                        MessageFactory.Text($"No search query was provided. Please provide a search query after the search command: `search <search terms here>`."), cancellationToken);
                    return;
                }

                //-- Set up the request.
                var requestBody = new SearchFromMSTeamsRequest()
                {
                    Email = email.account.Email,
                    Query = query
                };

                //-- Send the search string to the API.
                using var client = MakeRequest("Integrations/MSTeams/Search", out var request);

                //-- Add the request body.
                request.AddJsonBody(requestBody);

                //-- Execute the request.
                var response = await client.ExecutePostAsync(request, cancellationToken);
                if (response.StatusCode == System.Net.HttpStatusCode.OK)
                {
                    //-- Report the problem to the user.
                    SearchFromMSTeamsResults results = null;
                    if (response.Content != null)
                    {
                        results = JsonSerializer.Deserialize<SearchFromMSTeamsResults>(response.Content);
                    }

                    if (results == null || results.Results == null || results.Results.Length == 0)
                    {
                        //-- Return the message that no results were attained.
                        await turnContext.SendActivityAsync(
                            MessageFactory.Text($"No search results were found for '{query}'..."), cancellationToken);
                        return;
                    }

                    var adaptiveCard = new AdaptiveCard("1.2")
                    {
                        Body = new List<AdaptiveElement>()
                        {
                            new AdaptiveTextBlock
                            {
                                Text = $"Top search results for '{query}'...",
                                Weight = AdaptiveTextWeight.Bolder,
                                Size = AdaptiveTextSize.Large,
                            }
                        }
                    };

                    foreach (var res in results.Results) 
                    { 
                        adaptiveCard.Body.Add(new AdaptiveTextBlock
                        {
                            Text = $"**{res.Table}**: {res.Name} [(Open in Essembi)]({res.Url})",
                            Wrap = true
                        });
                    }

                    var attachment = new Attachment
                    {
                        ContentType = "application/vnd.microsoft.card.adaptive",
                        Content = adaptiveCard
                    };

                    await turnContext.SendActivityAsync(MessageFactory.Attachment(attachment), cancellationToken);
                }
                else
                {
                    //-- Return an error message.
                    await turnContext.SendActivityAsync(
                        MessageFactory.Text($"Authentication with Essembi failed. If you do not have an Essembi account you can register at essembi.com."), cancellationToken);
                }

            }
            else if (lowerText.Contains("help"))
            {
                var adaptiveCard = new AdaptiveCard("1.0")
                {
                    Body = new List<AdaptiveElement>()
                    {
                        new AdaptiveTextBlock
                        {
                            Text = "Help",
                            Weight = AdaptiveTextWeight.Bolder,
                            Size = AdaptiveTextSize.Large,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = "I am Essembi for Microsoft Teams. I can create tickets for you in Essembi easily from Teams chats and channels.",
                            Wrap = true
                        },
                        new AdaptiveTextBlock
                        {
                            Text = "I can be found under the 'actions and apps' button when composing new messages and under the 'more actions' menu on existing messages. I respond to the 'search', 'help' and 'documentation' commands.",
                            Wrap = true
                        }
                    }
                };

                var attachment = new Attachment
                {
                    ContentType = "application/vnd.microsoft.card.adaptive",
                    Content = adaptiveCard
                };

                await turnContext.SendActivityAsync(MessageFactory.Attachment(attachment), cancellationToken);
            }
            else if (lowerText.Contains("doc"))
            {
                var adaptiveCard = new AdaptiveCard("1.0")
                {
                    Body = new List<AdaptiveElement>()
                    {
                        new AdaptiveTextBlock
                        {
                            Text = "Documentation",
                            Weight = AdaptiveTextWeight.Bolder,
                            Size = AdaptiveTextSize.Large,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = "Learn more about Essembi for Microsoft Teams by navigating to our help documentation with the button below.",
                            Wrap = true
                        }
                    },
                    Actions = new List<AdaptiveAction>()
                    {
                        new AdaptiveOpenUrlAction
                        {
                            Title = "Learn more on essembi.com",
                            Url = new Uri("https://essembi.com/blogs/help/microsoft-teams-integration")
                        }
                    }
                };

                var attachment = new Attachment
                {
                    ContentType = "application/vnd.microsoft.card.adaptive",
                    Content = adaptiveCard
                };

                await turnContext.SendActivityAsync(MessageFactory.Attachment(attachment), cancellationToken);
            }
            else
            {
                //-- Suggesst the help or documentation actions.
                var heroCard = new HeroCard
                {
                    Title = "Unknown Command",
                    Subtitle = "I'm sorry. I do not recognize this command. Here are the commands that I currently recognize.",
                    Buttons = new List<CardAction>
                    {
                        new CardAction
                        {
                            Type = ActionTypes.MessageBack,
                            Title = "Documentation",
                            Text = "doc"
                        },
                        new CardAction
                        {
                            Type = ActionTypes.MessageBack,
                            Title = "Help",
                            Text = "help"
                        },
                        new CardAction
                        {
                            Type = ActionTypes.MessageBack,
                            Title = "Search",
                            Text = $"search {text}"
                        },
                    }
                };

                await turnContext.SendActivityAsync(MessageFactory.Attachment(heroCard.ToAttachment()), cancellationToken);
            }
        }

        #endregion

        #region Welcome Message

        async Task SendWelcomeCard(ITurnContext turnContext, CancellationToken cancellationToken)
        {
            var adaptiveCard = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>()
                {
                    new AdaptiveTextBlock
                    {
                        Text = "Hello!",
                        Weight = AdaptiveTextWeight.Bolder,
                        Size = AdaptiveTextSize.Large,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = "I am Essembi for Microsoft Teams. I can create tickets for you in Essembi easily from Teams chats and channels.",
                        Wrap = true
                    },
                    new AdaptiveTextBlock
                    {
                        Text = "I can be found under the 'actions and apps' button when composing new messages and under the 'more actions' menu on existing messages. I respond to the 'help' and 'documentation' commands.",
                        Wrap = true
                    },
                    new AdaptiveTextBlock
                    {
                        Text = "To use Essembi for Microsoft Teams, be sure to enable the the integration in Essembi. For more information on this, see our documentation on essembi.com.",
                        Wrap = true
                    }
                },
                Actions = new List<AdaptiveAction>()
                {
                    new AdaptiveOpenUrlAction
                    {
                        Title = "Learn more on essembi.com",
                        Url = new Uri("https://essembi.com/blogs/help/microsoft-teams-integration")
                    }
                }
            };

            var attachment = new Attachment
            {
                ContentType = "application/vnd.microsoft.card.adaptive",
                Content = adaptiveCard
            };

            await turnContext.SendActivityAsync(MessageFactory.Attachment(attachment), cancellationToken);
        }

        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            foreach (var member in turnContext.Activity.MembersAdded)
            {
                if (member.Id != turnContext.Activity.Recipient.Id)
                {
                    await SendWelcomeCard(turnContext, cancellationToken);
                }
            }
        }

        #endregion

        #region Utility Functions

        static MessagingExtensionActionResponse ErrorMessage(string message)
        {
            return new MessagingExtensionActionResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo
                    {
                        Card = new Attachment
                        {
                            ContentType = "application/vnd.microsoft.card.adaptive",
                            Content = new AdaptiveCard("1.0")
                            {
                                Body = new List<AdaptiveElement>()
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = message,
                                        Wrap = true
                                    }
                                },
                                Actions = new List<AdaptiveAction>()
                                {
                                    new AdaptiveOpenUrlAction
                                    {
                                        Title = "Contact Support",
                                        Url = new Uri("https://essembi.com/pages/support")
                                    }
                                }
                            }
                        },
                        Height = "small",
                        Width = "small",
                        Title = "An Issue has Occurred",
                    },
                },
            };
        }

        RestClient MakeRequest(string path, out RestRequest request)
        {
            //-- Get the web service URL.
            var serviceUrl = $"{Startup.ServiceBaseUrl}/{path}";

            //-- Instantiate the request.
            var options = new RestClientOptions(serviceUrl)
            {
                ThrowOnAnyError = false
            };

            var client = new RestClient(options);
            request = new RestRequest();

            //-- Add the authentication key.
            request.AddHeader("Authorization", $"Bearer {_teamsKey}");

            return client;
        }

        static MessagingExtensionActionResponse ShowTicketEntryWindow(string email, string subject, string body, EssembiApp app)
        {
            if(string.IsNullOrEmpty(subject) && !string.IsNullOrEmpty(body))
            {
                //-- Use the first line of the body as the subject.
                subject = body.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
            }

            //-- Create the dynamic form based on the setup in Essembi.
            var card = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>(),
                Actions = new List<AdaptiveAction>()
                {
                    new AdaptiveSubmitAction
                    {
                        Title = "Submit",
                        Type = "Action.Submit",
                        Data = new Dictionary<string, object>
                        {
                            { "email", email },
                            { "appId", app.Id },
                            { "tableId", app.TableId }
                        }
                    }
                }
            };

            //-- Add the short text fields (summary).
            var sortedFields = app.Fields.OrderBy(f => f.Name);
            foreach (var field in sortedFields.Where(f => f.Type == "shortText"))
            {
                card.Body.Add(new AdaptiveTextInput
                {
                    Label = field.Name,
                    IsRequired = field.Required,
                    Id = field.Id.ToString(),
                    Value = subject ?? string.Empty
                });

                //-- Use the subject on the first short text field (will be ticket name by default).
                subject = null;
            }

            //-- Add the categorical fields (category, type, priority, etc).
            foreach (var field in sortedFields.Where(f => f.Type == "record"))
            {
                card.Body.Add(new AdaptiveChoiceSetInput
                {
                    Label = field.Name,
                    IsRequired = field.Required,
                    Id = field.Id.ToString(),
                    Choices = field.Values.Select(v => new AdaptiveChoice
                    {
                        Title = v.Name,
                        Value = v.Id.ToString()
                    }).ToList()
                });
            }

            //-- Add the long text fields (description / story).
            foreach (var field in sortedFields.Where(f => f.Type == "longText"))
            {
                card.Body.Add(new AdaptiveTextInput
                {
                    Label = field.Name,
                    IsRequired = field.Required,
                    IsMultiline = true,
                    Id = field.Id.ToString(),
                    Value = body ?? string.Empty
                });

                //-- Use the body on the first long text field (will be ticket description by default).
                body = null;
            }

            //-- Return the form.
            return new MessagingExtensionActionResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo
                    {
                        Card = new Attachment
                        {
                            ContentType = "application/vnd.microsoft.card.adaptive",
                            Content = card
                        },
                        Height = card.Body.Count <= 5 ? "medium" : "large",
                        Width = "medium",
                        Title = "Create a Ticket in Essembi",
                    },
                },
            };
        }

        static Attachment GetAdaptiveCardAttachmentFromFile(string fileName)
        {
            //-- Read the card json and create attachment.
            string[] paths = { ".", "Resources", fileName };
            var adaptiveCardJson = File.ReadAllText(Path.Combine(paths));

            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = "application/vnd.microsoft.card.adaptive",
                Content = Newtonsoft.Json.JsonConvert.DeserializeObject(adaptiveCardJson),
            };
            return adaptiveCardAttachment;
        }

        #endregion
    }
}
