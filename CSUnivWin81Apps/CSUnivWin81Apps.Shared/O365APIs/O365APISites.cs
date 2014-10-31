using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using SampleModel.SharePoint;
using System.Net.Http.Headers;
using Windows.UI.Popups;
using Windows.Storage.FileProperties;

// Sample data structures for deserializing JSON data returned by SharePoint 
namespace SampleModel.SharePoint
{
    public class SharePointFileItem
    {
        public string FileLeafRef { get; set; }
        public SharePointFile File { get; set; }
        public DateTime Modified { get; set; }
        public int OData__ModerationStatus { get; set; }
        public SharePointFileMeta __metadata { get; set; }
    }

    public class SharePointFile
    {
        public SharePointUser Author { get; set; }
    }

    public class SharePointUser
    {
        public string Title { get; set; }
    }

    public class SharePointFileMeta
    {
        public string uri { get; set; }
    }

    public class ResultClass
    {
        public string itemUri { get; set; }
        public string Title { get; set; }
        public string ApprovalStatus { get; set; }
        public string TimeLastModified { get; set; }
        public string Author { get; set; }
        public object Thumbnail { get; set; }
    }
}

namespace CSUnivWin81Apps.O365APIs
{
    public class O365APISites
    {
        #region globals
        private static HttpClient httpClient = new HttpClient();

        //
        // The Client ID is used by the application to uniquely identify itself to Azure AD.
        // The Tenant is the name of the Azure AD tenant in which this application is registered.
        // The AAD Instance is the instance of Azure, for example public Azure or Azure China.
        // The Authority is the sign-in URL of the tenant.
        //

        const string aadInstance = "https://login.windows.net/{0}";
        const string tenant = "<tenant>.onmicrosoft.com";
        const string clientId = "client id from Azure AD";


        static string authority = String.Format(CultureInfo.InvariantCulture, aadInstance, tenant);

        //
        // To authenticate to the To Do list service, the client needs to know the service's App ID URI.
        // To contact the To Do list service we need it's URL as well.
        //
        const string todoListResourceId = "https://<tenant>.sharepoint.com/";
        const string todoListBaseAddress = "https://<tenant>.sharepoint.com";
        #endregion


        #region O365 _api Functions
        //Approve Document
        public static async Task<HttpResponseMessage> updateDocApproval(AuthenticationResult result, AuthenticationContext authContext, IList<ResultClass> theItem)
        {
            //TODO: need to figure out the batch version of this, right now for demo just performed the single file update
            //although I send over all items selected, I only process the first in the list.  Wanted to figure out how to do 
            //batch b/c that would be much more efficient obvi than looping thru the list.
            if (result.Status == AuthenticationStatus.Success)
            {

                // Add the access token to the Authorization Header of the call to the To Do list service, and call the service.
                //
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
                httpClient.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");
                httpClient.DefaultRequestHeaders.Add("X-HTTP-Method", "MERGE");
                httpClient.DefaultRequestHeaders.Add("IF-MATCH", "*");


                string metadataString = "{ '__metadata': { 'type': 'SP.Data.Shared_x0020_DocumentsItem' }, 'OData__ModerationStatus': 0 }";
                HttpContent content = new StringContent(metadataString);
                content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                content.Headers.ContentType.Parameters.Add(new NameValueHeaderValue("odata", "verbose"));
                MessageDialog dialog = new MessageDialog("You are about to Approve the following doc - " + theItem[0].Title);
                await dialog.ShowAsync();
                // Call the web api
                return await httpClient.PostAsync(theItem[0].itemUri, content);

            }
            else
            {
                MessageDialog gendialog = new MessageDialog(string.Format("If the error continues, please contact your administrator.\n\nError: {0}\n\nError Description:\n\n{1}", result.Error, result.ErrorDescription), "Sorry, an error occurred while signing you in.");
                await gendialog.ShowAsync();
                return null;
            }
        }

        // Retrieve the user's To Do list.
        public static async Task<HttpResponseMessage> GetDocsForApproval(AuthenticationResult result)
        {

            if (result.Status == AuthenticationStatus.Success)
            {
                //
                // Add the access token to the Authorization Header of the call to the To Do list service, and call the service.
                //
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
                httpClient.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");
               //Original Odata query with the pending approval filter
                // return await httpClient.GetAsync(todoListBaseAddress + "/_api/web/lists/getbytitle('Documents')/items?$select=FileLeafRef,File/Author/Title,OData__ModerationStatus,Modified,File/uri&$expand=File/Author&$filter=OData__ModerationStatus eq 2");
                return await httpClient.GetAsync(todoListBaseAddress + "/_api/web/lists/getbytitle('Documents')/items?$select=FileLeafRef,File/Author/Title,OData__ModerationStatus,Modified,File/uri&$expand=File/Author");
            }
            else
            {
                MessageDialog dialog = new MessageDialog(string.Format("If the error continues, please contact your administrator.\n\nError: {0}\n\nError Description:\n\n{1}", result.Error, result.ErrorDescription), "Sorry, an error occurred while signing you in.");
                await dialog.ShowAsync();
                return null;
            }
        }
        #endregion

        #region Helper Functions
        //Helper method to transform datetime from SP to readable format
        public static string ToLocalTimeString(DateTime dateTime)
        {
            return dateTime.ToLocalTime().ToString("g", CultureInfo.CurrentCulture);
        }

        //Helper method to transform 
        public static string ToApprovalStatusString(int modStatusCode)
        {
            switch (modStatusCode)
            {
                case 0:
                    return "Approved";
                case 1:
                    return "Rejected";
                case 2:
                    return "Pending";
                default:
                    return "StatusNotFound";
            }
        }
        #endregion
    }
}
