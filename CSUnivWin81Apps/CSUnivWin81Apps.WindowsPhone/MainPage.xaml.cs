using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;
using SampleModel.SharePoint;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Windows.UI.Popups;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Windows.ApplicationModel.Activation;
using Windows.Storage;


// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

namespace CSUnivWin81Apps
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page, IWebAuthenticationContinuable
    {
        #region init

        //
        // The Client ID is used by the application to uniquely identify itself to Azure AD.
        // The Tenant is the name of the Azure AD tenant in which this application is registered.
        // The AAD Instance is the instance of Azure, for example public Azure or Azure China.
        // The Authority is the sign-in URL of the tenant.
        //

        const string aadInstance = "https://login.windows.net/{0}";
        const string tenant = "<tenant>.onmicrosoft.com";
        const string clientId = "your client Id from Azure AD";


        static string authority = String.Format(CultureInfo.InvariantCulture, aadInstance, tenant);

        //
        // To authenticate to the To Do list service, the client needs to know the service's App ID URI.
        // To contact the To Do list service we need it's URL as well.
        //
        const string todoListResourceId = "https://<tenant>.sharepoint.com/";
        const string todoListBaseAddress = "https://<tenant>.sharepoint.com";

        private AuthenticationContext authContext = null;
        private Uri redirectURI = null;

        #endregion

        public MainPage()
        {
            this.InitializeComponent();

            this.NavigationCacheMode = NavigationCacheMode.Required;
            //
            // Every Windows Store application has a unique URI.
            // Windows ensures that only this application will receive messages sent to this URI.
            // ADAL uses this URI as the application's redirect URI to receive OAuth responses.
            // 
            // To determine this application's redirect URI, which is necessary when registering the app
            //      in AAD, set a breakpoint on the next line, run the app, and copy the string value of the URI.
            //      This is the only purposes of this line of code, it has no functional purpose in the application.
            //
            redirectURI = Windows.Security.Authentication.Web.WebAuthenticationBroker.GetCurrentApplicationCallbackUri();

            // ADAL for Windows Phone 8.1 builds AuthenticationContext instances throuhg a factory, which performs authority validation at creation time
            authContext = AuthenticationContext.CreateAsync(authority).GetResults();
            //AuthAppandCallFunc("GetDocsForApproval");
        }

        /// <summary>
        /// Invoked when this page is about to be displayed in a Frame.
        /// </summary>
        /// <param name="e">Event data that describes how this page was reached.
        /// This parameter is typically used to configure the page.</param>
        protected override void OnNavigatedTo(NavigationEventArgs e)
        {
            // TODO: Prepare page for display here.

            // TODO: If your application contains multiple pages, ensure that you are
            // handling the hardware Back button by registering for the
            // Windows.Phone.UI.Input.HardwareButtons.BackPressed event.
            // If you are using the NavigationHelper provided by some templates,
            // this event is handled for you.
        }

        #region O365 APIs Interface implmentation for IWebAuthenticationContinuable

        // This method is automatically invoked when the application is reactivated after an authentication interaction throuhg WebAuthenticationBroker.        
        public async void ContinueWebAuthentication(WebAuthenticationBrokerContinuationEventArgs args)
        {
            // pass the authentication interaction results to ADAL, which will conclude the token acquisition operation and invoke the callback specified in AcquireTokenAndContinue.
            await authContext.ContinueAcquireTokenAsync(args);
        }
        #endregion

        /// <summary>
        /// Called from tapped event to update document content approval status
        /// </summary>
        /// <param name="result"></param>
        public async void updateDocApproval(AuthenticationResult result)
        {

            //Lets get the value to update
            var theItem = (IList<ResultClass>)itemListView.SelectedItems.Cast<ResultClass>().ToList();
            var response = await O365APIs.O365APISites.updateDocApproval(result, authContext, theItem);
            
            //if null then catastrophic error and return processing back to client, already reported to user
            if (response == null)
                return;
            if (response.IsSuccessStatusCode)
            {
                MessageDialog successdialog = new MessageDialog("Success! Doc - " + theItem[0].Title + " - is Approved.");
                await successdialog.ShowAsync();
                // Read the response as a Json Array and databind to the GridView to display remaining docs for approval
                GetDocsForApproval(result);
            }
            else
            {
                if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {
                    // If the To Do list service returns access denied, clear the token cache and have the user sign-in again.
                    MessageDialog Unauthdialog = new MessageDialog("Sorry, you don't have access to approve these documents.  Please sign-in again.");
                    await Unauthdialog.ShowAsync();
                    authContext.TokenCache.Clear();
                }
                else
                {
                    MessageDialog errdialog = new MessageDialog("Sorry, an error occurred accessing your document library.  Please try again.");
                    await errdialog.ShowAsync();
                }
            }
        }

        /// <summary>
        /// Calls _api in Office 365 to retrieve default documents list in identified library
        /// </summary>
        /// <param name="result"></param>
        public async void GetDocsForApproval(AuthenticationResult result)
        {
            var response = await O365APIs.O365APISites.GetDocsForApproval(result);
            //exit code b/c there was catastrophic error and already reported it to user
            if (response == null)
                return;
            if (response.IsSuccessStatusCode)
            {
                // Read the response and deserialize the data: 
                string responseString = await response.Content.ReadAsStringAsync();

                var todoArray = JObject.Parse(responseString)["d"]["results"].ToObject<SampleModel.SharePoint.SharePointFileItem[]>();
                //if (todoArray.Count() == 0)
                //{

                //    Message1.Text = "No documents for Approval";
                //}

                itemListView.ItemsSource = from todo in todoArray
                                       select new ResultClass
                                       {
                                           Title = todo.FileLeafRef,
                                           Author = "Author: " + todo.File.Author.Title,
                                           ApprovalStatus = "Approval: " + O365APIs.O365APISites.ToApprovalStatusString(todo.OData__ModerationStatus),
                                           TimeLastModified = "Last Modified: " + O365APIs.O365APISites.ToLocalTimeString(todo.Modified),
                                           itemUri = todo.__metadata.uri
                                       };


                itemListViewApprove.ItemsSource = from todo in todoArray
                                                  where todo.OData__ModerationStatus == 2
                                                  select new ResultClass
                                                  {
                                                      Title = todo.FileLeafRef,
                                                      Author = "Author: " + todo.File.Author.Title,
                                                      ApprovalStatus = "Approval: " + O365APIs.O365APISites.ToApprovalStatusString(todo.OData__ModerationStatus),
                                                      TimeLastModified = "Last Modified: " + O365APIs.O365APISites.ToLocalTimeString(todo.Modified),
                                                      itemUri = todo.__metadata.uri
                                                  };
            }
            else
            {
                if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {
                    // If the To Do list service returns access denied, clear the token cache and have the user sign-in again.
                    MessageDialog dialog = new MessageDialog("Sorry, you don't have access to the To Do Service.  Please sign-in again.");
                    await dialog.ShowAsync();
                    authContext.TokenCache.Clear();
                }
                else
                {
                    MessageDialog dialog = new MessageDialog("Sorry, an error occurred accessing your To Do list.  Please try again.");
                    await dialog.ShowAsync();
                }
            }
        }

        //authenticate if needed and call function that is sent into method
        private async void AuthAppandCallFunc(string callingFunc)
        {
            // Try to get a token without triggering any user prompt. 
            // ADAL will check whether the requested token is in the cache or can be obtained without user itneraction (e.g. via a refresh token).
            AuthenticationResult result = await authContext.AcquireTokenSilentAsync(todoListResourceId, clientId);
            if (result != null && result.Status == AuthenticationStatus.Success)
            {
                if (callingFunc == "GetDocsForApproval")
                {
                    // A token was successfully retrieved. Get the To Do list for the current user
                    GetDocsForApproval(result);
                }
                else if (callingFunc == "UpdateDocApproval")
                {
                    updateDocApproval(result);
                }
            }
            else
            {
                if (callingFunc == "GetDocsForApproval")
                {
                    // Acquiring a token without user interaction was not possible. 
                    // Trigger an authentication experience and specify that once a token has been obtained the GetDocsForApproval method should be called
                    authContext.AcquireTokenAndContinue(todoListResourceId, clientId, redirectURI, GetDocsForApproval);
                }
                else if (callingFunc == "UpdateDocApproval")
                {
                    // Acquiring a token without user interaction was not possible. 
                    // Trigger an authentication experience and specify that once a token has been obtained the GetTodoList method should be called
                    authContext.AcquireTokenAndContinue(todoListResourceId, clientId, redirectURI, updateDocApproval);
                }
            }
        }

        private async void RefreshAppBarButton_Click(object sender, RoutedEventArgs e)
        {
            //AuthAppandCallFunc("GetDocsForApproval");
            // Try to get a token without triggering any user prompt. 
            // ADAL will check whether the requested token is in the cache or can be obtained without user itneraction (e.g. via a refresh token).
            AuthenticationResult result = await authContext.AcquireTokenSilentAsync(todoListResourceId, clientId);
            if (result != null && result.Status == AuthenticationStatus.Success)
            {
                // A token was successfully retrieved. Get the To Do list for the current user
                GetDocsForApproval(result);
            }
            else
            {
                // Acquiring a token without user interaction was not possible. 
                // Trigger an authentication experience and specify that once a token has been obtained the GetTodoList method should be called
                authContext.AcquireTokenAndContinue(todoListResourceId, clientId, redirectURI, GetDocsForApproval);
            }
        }

        private void RemoveAppBarButton_Click(object sender, RoutedEventArgs e)
        {
            AuthAppandCallFunc("UpdateDocApproval");
        }

        private void DeleteConfirmation_Click(object sender, RoutedEventArgs e)
        {

        }

        private void itemListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (itemListView.SelectedItems.Count > 0)
            {
                itemListView.SelectionMode = ListViewSelectionMode.Multiple;
            }
            else
            {
                itemListView.SelectionMode = ListViewSelectionMode.Single;
            }
        }

        private void itemListViewApproved_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (itemListViewApprove.SelectedItems.Count > 0)
            {
                itemListViewApprove.SelectionMode = ListViewSelectionMode.Multiple;
            }
            else
            {
                itemListViewApprove.SelectionMode = ListViewSelectionMode.Single;
            }
        }
    }
}
