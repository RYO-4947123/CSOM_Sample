using Microsoft.SharePoint.Client;
using PnP.Core.Model;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Demo.WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly IPnPContextFactory pnpContextFactory;

        public MainWindow(IPnPContextFactory pnpFactory)
        {
            this.pnpContextFactory = pnpFactory;

            InitializeComponent();
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

        internal async Task SiteInfoAsync()
        {
            using (var context = await pnpContextFactory.CreateAsync("DemoSite"))
            {
                var accessToken =  await context.AuthenticationProvider.GetAccessTokenAsync(context.Uri);
                //Console.WriteLine($"Token: {accessToken}");
                var CSOMContext = GetContext(context.Uri, accessToken);
                using(CSOMContext) {
                    Web CSOMWeb = CSOMContext.Web;
                    CSOMContext.Load(CSOMWeb, w => w.Title, w => w.Description);
                    CSOMContext.ExecuteQuery();
                    Console.WriteLine(CSOMWeb.Title);
                }
            }
        }

        private async void btnSite_Click(object sender, RoutedEventArgs e)
        {
            await SiteInfoAsync();
        }
    }
}
