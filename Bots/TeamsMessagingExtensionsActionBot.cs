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
                    Title = "View Ticket",
                    Value = successMessage.Url
                },
                Buttons = new List<CardAction>
                {
                    new CardAction
                    {
                        Type = ActionTypes.OpenUrl,
                        Title = "View Ticket",
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

            TeamsChannelAccount member;
            try
            {
                //-- Get member information.
                member = await TeamsInfo.GetMemberAsync(turnContext, turnContext.Activity.From.Id, cancellationToken);
            }
            catch (ErrorResponseException ex)
            {
                if (ex.Body.Error.Code == "BotNotInConversationRoster" || ex.Response.StatusCode == System.Net.HttpStatusCode.Forbidden)
                {
                    //-- The bot is not in the conversation roster.
                    return new MessagingExtensionActionResponse
                    {
                        Task = new TaskModuleContinueResponse
                        {
                            Value = new TaskModuleTaskInfo
                            {
                                Card = GetAdaptiveCardAttachmentFromFile("justintimeinstallation.json"),
                                Height = 200,
                                Width = 400,
                                Title = "Adaptive Card - App Installation",
                            },
                        },
                    };
                }

                //-- It's a different error.
                throw;
            }

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

            string subject = null, body = null;
            if (action.MessagePayload != null)
            {
                subject = action.MessagePayload.Subject;
                body = action.MessagePayload.Body?.Content;

                if (body != null)
                {                    
                    //-- Remove the HTML tags.
                    var doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(body);
                    body = doc.DocumentNode.InnerText;
                }
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
