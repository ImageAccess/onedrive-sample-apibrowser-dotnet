using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Diagnostics;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace OneDriveApiBrowser
{
    public class AuthenticationHelper
    {
        // The Client ID is used by the application to uniquely identify itself to the v2.0 authentication endpoint.
        private static string clientId = FormBrowser.MsaClientId;

        private static string redirectURI = FormBrowser.MsaReturnUrl;

        public static string[] Scopes = { "Files.ReadWrite Files.ReadWrite.All" };

        public static PublicClientApplicationBuilder IdentityClientAppBuilder = PublicClientApplicationBuilder.Create (clientId)
            .WithRedirectUri (redirectURI)
            .WithAuthority(AzureCloudInstance.AzurePublic, "common");

        public static IPublicClientApplication IdentityClientApp = IdentityClientAppBuilder.Build();

        public static IAccount Account;
        public static string TokenForUser = null;
        public static DateTimeOffset Expiration;

        private static GraphServiceClient graphClient = null;

        // Get an access token for the given context and resourceId. An attempt is first made to
        // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.
        public static GraphServiceClient GetAuthenticatedClient ()
        {
            if (graphClient == null)
            {
                // Create Microsoft Graph client.
                try
                {
                    graphClient = new GraphServiceClient
                    ("https://graph.microsoft.com/v1.0",
                        new DelegateAuthenticationProvider
                        (
                            async (requestMessage) =>
                            {
                                var token = await GetTokenForUserAsync();
                                requestMessage.Headers.Authorization = new AuthenticationHeaderValue ("bearer", token);
                            }
                        )
                    );
                    return graphClient;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine ("Could not create a graph client: " + ex.Message);
                }
            }
            return graphClient;
        }

        /// <summary>
        /// Get Token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static async Task<string> GetTokenForUserAsync ()
        {
            AuthenticationResult authResult;
            try
            {
                authResult = await IdentityClientApp.AcquireTokenSilent (Scopes, Account).ExecuteAsync ();
                TokenForUser = authResult.AccessToken;
            }
            catch (Exception)
            {
                if (TokenForUser == null || Expiration <= DateTimeOffset.UtcNow.AddMinutes (5))
                {
                    authResult = await IdentityClientApp.AcquireTokenInteractive (Scopes.AsEnumerable<string> ()).ExecuteAsync ();
                    TokenForUser = authResult.AccessToken;
                    Expiration = authResult.ExpiresOn;
                }
            }
            return TokenForUser;
        }

        /// <summary>
        /// Signs the user out of the service.
        /// </summary>
        public static void SignOut ()
        {
            IdentityClientApp.GetAccountsAsync ();

            foreach (var user in IdentityClientApp.GetAccountsAsync ().Result)
            {
                //user.SignOut ();
            }
            graphClient = null;
            TokenForUser = null;
        }
    }
}