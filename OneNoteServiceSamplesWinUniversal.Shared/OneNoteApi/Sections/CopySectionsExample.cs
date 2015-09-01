//*********************************************************
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

namespace OneNoteServiceSamplesWinUniversal.OneNoteApi.Sections
{
	/// <summary>
    /// Class to show examples of copying sections via HTTP POST to the CopyToNotebook and CopyToSectionGroup actions. 
    /// If successful, returns a 202 Accepted status code, a Location header with the section's endpoint URL, and a Microsoft.OneNote.Api.CopyStatusModel object.
	/// </summary>
	/// <remarks>
	/// NOTE: These actions require the ID of the section that you want to copy and the ID of the notebook or section group  
    /// that you want to copy the section to. 
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
    ///  var createMessage = new HttpRequestMessage(HttpMethod.Post, "https://www.onenote.com/api/beta/me/notes/sections/{id}/Microsoft.OneNote.Api.CopyToNotebook")
	///  {
	///      Content = new StringContent("{ id : 'section-id', renameAs: 'New Section Name' }", System.Text.Encoding.UTF8, "application/json")
	///  };
	///  HttpResponseMessage response = await client.SendAsync(createMessage);
	/// </code>
	public static class CopySectionsExample
    {
        #region Examples of POST https://www.onenote.com/api/beta/me/notes/sections/{id}/Microsoft.OneNote.Api.CopyToNotebook

        /// <summary>
		/// Copy a section to a notebook
		/// </summary>
        /// <param name="debug">Run the code under the debugger</param>
        /// <param name="sectionId">Id of the section to copy</param>
        /// <param name="notebookId">Id of the destination notebook</param>
        /// <param name="newSectionName">Name for the new section. Must be unique in the destination notebook.</param>
		/// <param name="provider">Live Connect or Azure AD</param>
        /// <param name="apiRoute">v1.0 or beta path</param>
		/// <remarks></remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<ApiBaseResponse> CopySectionToNotebook(bool debug, string sectionId, string notebookId, string newSectionName, AuthProvider provider, string apiRoute)
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

			// Prepare an HTTP POST request to the CopyToNotebook endpoint of the target section.
            // The request body content type is application/json. 
            // Request body parameters: 
            //   - id: The ID of the notebook to copy to. Required for all CopyToNotebook calls. 
            //   - siteCollectionId and siteId: The SharePoint site to copy the section to. Required to copy to a site.
            //   - groupId: The ID of the group to copy the section to. Required to copy to a group.
            //   - renameAs: The name for the copy of the section. Returns 409 (Conflict) if a section with the same name exists 
            //         in the destination notebook. If omitted, defaults to the name of the existing section.
            var createMessage = new HttpRequestMessage(HttpMethod.Post, apiRoute + "sections/" + sectionId + "/Microsoft.OneNote.Api.CopyToNotebook")
			{
				Content = new StringContent("{ id : '" + notebookId + "',  renameAs : '" + newSectionName + "' }", Encoding.UTF8, "application/json")
			};
			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await HttpUtils.TranslateResponse(response);
		}

		#endregion

        #region Examples of POST https://www.onenote.com/api/beta/me/notes/sections/{id}/Microsoft.OneNote.Api.CopyToSectionGroup

        /// <summary>
        /// Copy a section to a section group
        /// </summary>
        /// <param name="debug">Run the code under the debugger</param>
        /// <param name="sectionId">Id of the section to copy</param>
        /// <param name="notebookId">Id of the destination section group</param>
        /// <param name="newSectionName">Name for the new section. Must be unique in the destination section group.</param>
        /// <param name="provider">Live Connect or Azure AD</param>
        /// <param name="apiRoute">v1.0 or beta path</param>
        /// <remarks></remarks>
        /// <returns>The converted HTTP response message</returns>
        public static async Task<ApiBaseResponse> CopySectionToSectionGroup(bool debug, string sectionId, string sectionGroupId, string newSectionName, AuthProvider provider, string apiRoute)
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

            // Prepare an HTTP POST request to the CopyToSectionGroup endpoint of the target section.
            // The request body content type is application/json. 
            // Request body parameters: 
            //   - id: The ID of the section group to copy to. Required for all CopyToSectionGroup calls. 
            //   - siteCollectionId and siteId: The SharePoint site to copy the section to. Required to copy to a site.
            //   - groupId: The ID of the group to copy to the section to. Required to copy to a group.
            //   - renameAs: The name for the copy of the section. Returns 409 (Conflict) if a section with the same name exists 
            //         in the destination section group. If omitted, defaults to the name of the existing section.
            var createMessage = new HttpRequestMessage(HttpMethod.Post, apiRoute + "sections/" + sectionId + "/Microsoft.OneNote.Api.CopyToSectionGroup")
            {
                Content = new StringContent("{ id : '" + sectionGroupId + "',  renameAs : '" + newSectionName + "' }", Encoding.UTF8, "application/json")
            };
            HttpResponseMessage response = await client.SendAsync(createMessage);

            return await HttpUtils.TranslateResponse(response);
        }

        #endregion
	}
}
