﻿//*********************************************************
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// Licensed under the Apache License, Version 2.0 (the ""License""); 
// you may not use this file except in compliance with the License. 
// You may obtain a copy of the License at 
// http://www.apache.org/licenses/LICENSE-2.0 
//
// THIS CODE IS PROVIDED ON AN  *AS IS* BASIS, WITHOUT 
// WARRANTIES OR CONDITIONS OF ANY KIND, EITHER EXPRESS 
// OR IMPLIED, INCLUDING WITHOUT LIMITATION ANY IMPLIED 
// WARRANTIES OR CONDITIONS OF TITLE, FITNESS FOR A PARTICULAR 
// PURPOSE, MERCHANTABLITY OR NON-INFRINGEMENT. 
//
// See the Apache Version 2.0 License for specific language 
// governing permissions and limitations under the License.
//*********************************************************
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace OneNoteServiceSamplesWinUniversal.OneNoteApi.Pages
{
	/// <summary>
    /// Class to show an example of copying notebooks via HTTP POST to the CopyToSection action. 
    /// If successful, returns a 202 Accepted status code, a Location header with the page's endpoint URL, and a Microsoft.OneNote.Api.CopyStatusModel object.
	/// </summary>
	/// <remarks>
	/// NOTE: This action requires the ID of the page that you want to copy and the ID of the section you want to copy the page to.
	/// NOTE: These examples do not include copying to a SharePoint site or group. 
	/// NOTE: It is not the goal of this code sample to produce well re-factored, elegant code.
	/// You may notice code blocks being duplicated in various places in this project.
	/// We have deliberately added these code blocks to allow anyone browsing the sample
	/// to easily view all related functionality in near proximity.
	/// </remarks>
	/// <code>
	///  var client = new HttpClient(); 
	///  client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
	///  client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await Auth.GetAuthToken());
    ///  var createMessage = new HttpRequestMessage(HttpMethod.Post, "https://www.onenote.com/api/beta/me/notes/pages/{id}/Microsoft.OneNote.Api.CopyToSection")
	///  {
    ///      Content = new StringContent("{ id : 'page-id', renameAs: 'New Page Name' }", System.Text.Encoding.UTF8, "application/json")
	///  };
	///  HttpResponseMessage response = await client.SendAsync(createMessage);
	/// </code>
	public static class CopyPagesExample
    {
        #region Examples of POST https://www.onenote.com/api/beta/me/notes/pages/{id}/Microsoft.OneNote.Api.CopyToSection

        /// <summary>
		/// Copy a page to a section
		/// </summary>
        /// <param name="debug">Run the code under the debugger</param>
        /// <param name="pageId">ID of the page to copy</param>
        /// <param name="pageId">ID of the destination section</param>
		/// <param name="newPageName">Name for the new page</param>
		/// <param name="provider">Live Connect or Azure AD</param>
        /// <param name="apiRoute">v1.0 or beta path</param>
		/// <remarks></remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<ApiBaseResponse> CopyPageToSection(bool debug, string pageId, string sectionId, string newPageName, AuthProvider provider, string apiRoute)
		{
			if (debug)
			{
				Debugger.Launch();
				Debugger.Break();
			}

			var client = new HttpClient();

			// Note: The API returns OneNote entities as JSON objects.
			client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

			// Not adding the Authentication header results in an unauthorized call and the API returns a 401 status code.
			client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer",
				await Auth.GetAuthToken(provider));

			// Prepare an HTTP POST request to the CopyToSection endpoint of the target page.
            // The request body content type is application/json. 
            // Request body parameters: 
            //   - id: The ID of the section to copy to. Required for all CopyToSection calls. 
            //   - siteCollectionId and siteId: The SharePoint site to copy the page to. Required to copy to a site.
            //   - groupId: The ID of the group to copy to the page to. Required to copy to a group.
            //   - renameAs: The name for the copy of the page. If omitted, defaults to the name of the existing page. //xx
            var createMessage = new HttpRequestMessage(HttpMethod.Post, apiRoute + "pages/" + pageId + "/Microsoft.OneNote.Api.CopyToSection")
			{
				Content = new StringContent("{ id : '" + sectionId + "', renameAs : '" + newPageName + "' }", Encoding.UTF8, "application/json")
			};
			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await HttpUtils.TranslateResponse(response);
		}

		#endregion

	}
}