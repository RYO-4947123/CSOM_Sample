using Microsoft.SharePoint.Client;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using PnP.Core.Auth;
using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Threading.Tasks;

namespace Consumer
{
    class Program
    {
        public static async Task Main(string[] args)
        {
            var host = Host.CreateDefaultBuilder()

            // Ensure you do consent to the PnP App when using another tenant (update below url to match your aad domain): 
            // https://login.microsoftonline.com/a830edad9050849523e17050400.onmicrosoft.com/adminconsent?client_id=31359c7f-bd7e-475c-86db-fdb8c937548e&state=12345&redirect_uri=https://www.pnp.com
            //.UseEnvironment("officedevpnp")
            .ConfigureLogging((hostingContext, logging) =>
            {
                logging.AddEventSourceLogger();
                logging.AddConsole();
            })
            .ConfigureServices((hostingContext, services) =>
            {
                // Read the custom configuration from the appsettings.<environment>.json file
                var customSettings = new CustomSettings();
                hostingContext.Configuration.Bind("CustomSettings", customSettings);

                // Create an instance of the Authentication Provider that uses Credential Manager
                //var authenticationProvider = new CredentialManagerAuthenticationProvider(
                //                customSettings.ClientId,
                //                customSettings.TenantId,
                //                customSettings.CredentialManager);                

                var authenticationProvider = new InteractiveAuthenticationProvider(
                                customSettings.ClientId,
                                customSettings.TenantId,
                                customSettings.RedirectUri);

                // Add the PnP Core SDK services
                services.AddPnPCore(options => {

                    // You can explicitly configure all the settings, or you can
                    // simply use the default values

                    //options.PnPContext.GraphFirst = true;
                    //options.PnPContext.GraphCanUseBeta = true;
                    //options.PnPContext.GraphAlwaysUseBeta = false;

                    //options.HttpRequests.UserAgent = "NONISV|SharePointPnP|PnPCoreSDK";
                    //options.HttpRequests.MicrosoftGraph = new PnPCoreHttpRequestsGraphOptions
                    //{
                    //    UseRetryAfterHeader = true,
                    //    MaxRetries = 10,
                    //    DelayInSeconds = 3,
                    //    UseIncrementalDelay = true,
                    //};
                    //options.HttpRequests.SharePointRest = new PnPCoreHttpRequestsSharePointRestOptions
                    //{
                    //    UseRetryAfterHeader = true,
                    //    MaxRetries = 10,
                    //    DelayInSeconds = 3,
                    //    UseIncrementalDelay = true,
                    //};

                    options.DefaultAuthenticationProvider = authenticationProvider;

                    options.Sites.Add("DemoSite",
                        new PnP.Core.Services.Builder.Configuration.PnPCoreSiteOptions
                        {
                            SiteUrl = customSettings.DemoSiteUrl,
                            AuthenticationProvider = authenticationProvider
                        });
                });
            })
            // Let the builder know we're running in a console
            .UseConsoleLifetime()
            // Add services to the container
            .Build();

            await host.StartAsync();

            using (var scope = host.Services.CreateScope())
            {
                var pnpContextFactory = scope.ServiceProvider.GetRequiredService<IPnPContextFactory>();

                #region Interactive GET's
                using (var context = await pnpContextFactory.CreateAsync("DemoSite"))
                {
                    var accessToken =  await context.AuthenticationProvider.GetAccessTokenAsync(context.Uri);
                    //Console.WriteLine($"Token: {accessToken}");
                    var CSOMContext = GetContext(context.Uri, accessToken);

                    using(CSOMContext) {
                        Web web = CSOMContext.Web;
                        CSOMContext.Load(web, w => w.Title, w => w.Description);
                        CSOMContext.ExecuteQuery();
                        Console.WriteLine(web.Title);
                    }
                }
                #endregion

            }

            host.Dispose();
        }

        private static ClientContext GetContext(Uri web, String accessToken) {
            var context = new ClientContext(web)
            {
                // Important to turn off FormDigestHandling when using access tokens
                //FormDigestHandlingEnabled = false
            };
            context.ExecutingWebRequest += (sender, e) =>
            {
                // Insert the access token in the request
                e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
            };

            return context;
        }
    }
}
