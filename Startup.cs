using Essembi.Integrations.Teams.Bots;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace Essembi.Integrations.Teams
{
    public class Startup
    {
        static string _overrideServiceBaseUrl;
        public static string ServiceBaseUrl
        {
            get
            {
#if DEBUG
                const string serviceBaseUrl = "https://localhost:7198";
#else
                const string serviceBaseUrl = "https://api.essembi.ai";
#endif

                return string.IsNullOrEmpty(_overrideServiceBaseUrl) ? serviceBaseUrl : _overrideServiceBaseUrl;
            }
        }

        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;

            _overrideServiceBaseUrl = configuration["ServiceBaseUrl"];

        }

        public IConfiguration Configuration { get; }

        /// <summary>
        /// This method gets called by the runtime. Use this method to add services to the container.
        /// </summary>
        /// <param name="services">The services.</param>
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddControllers();
            services.AddHttpClient();
            services.AddMvc();
            services.AddControllers().AddNewtonsoftJson(options =>
            {
                options.SerializerSettings.MaxDepth = HttpHelper.BotMessageSerializerSettings.MaxDepth;
            });
            services.AddRazorPages(c => c.RootDirectory = "/wwwroot").AddRazorPagesOptions(c => { });

            //-- Create the Bot Framework Authentication to be used with the Bot Adapter.
            services.AddSingleton<BotFrameworkAuthentication, ConfigurationBotFrameworkAuthentication>();

            //-- Create the Bot Adapter with error handling enabled.
            services.AddSingleton<IBotFrameworkHttpAdapter, AdapterWithErrorHandler>();

            //-- Add basic memory state storage to handle storing information between popups.
            services.AddSingleton<IStorage, MemoryStorage>();
            services.AddSingleton<UserState>();

            //-- Create the bot as a transient. In this case the ASP Controller is expecting an IBot.
            services.AddTransient<IBot, TeamsMessagingExtensionsActionBot>();
        }

        /// <summary>
        /// Configures the specified application.
        /// </summary>
        /// <param name="app">The application.</param>
        /// <param name="env">The env.</param>
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }

            app.UseDefaultFiles();
            app.UseStaticFiles();

            //-- Runs matching. An endpoint is selected and set on the HttpContext if a match is found.
            app.UseRouting();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapRazorPages();

                //-- Mapping of endpoints goes here:
                endpoints.MapControllers();
            });
        }
    }
}
