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
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace OneNoteServiceSamplesWinUniversal.OneNoteApi.Pages
{
	/// <summary>
	///     Class to show a selection of examples deleting pages via HTTP DELETE to the OneNote API
	///     - Delete a new page is represented via the DELETE HTTP verb.
    ///     For more info, see http://dev.onenote.com/docs#/reference/delete-pages.
	/// </summary>
	/// <remarks>
	///     NOTE: It is not the goal of this code sample to produce well re-factored, elegant code.
	///     You may notice code blocks being duplicated in various places in this project.
	///     We have deliberately added these code blocks to allow anyone browsing the sample
	///     to easily view all related functionality in near proximity.
	/// </remarks>
	/// <code>
	///   var client = new HttpClient();
	///   client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await Auth.GetAuthToken());
	/// 
	///   var deleteMessage = new HttpRequestMessage(HttpMethod.Delete, "https://www.onenote.com/api/v1.0/pages");
	///   HttpResponseMessage response = await client.SendAsync(deleteMessage);
	///  </code>
	public static class DeletePagesExample
	{
		#region Examples of DELETE https://www.onenote.com/api/v1.0/pages

		/// <summary>
		///     Delete a page.
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="pageId"></param>
		/// <param name="provider"></param>
		/// <param name="apiRoute"></param>
		/// <remarks>Delete a specified page by ID.</remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<ApiBaseResponse> DeletePage(bool debug, string pageId, AuthProvider provider, string apiRoute)
		{
			if (debug)
			{
				Debugger.Launch();
				Debugger.Break();
			}

			var client = new HttpClient();

			// Note: API only supports JSON response.
			client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

			// Not adding the Authentication header would produce an unauthorized call and the API will return a 401
			client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer",
				await Auth.GetAuthToken(provider));

			// Prepare an HTTP DELETE request to the Pages endpoint
			// The request body content type is text/html
			var deleteMessage = new HttpRequestMessage(HttpMethod.Delete, apiRoute + "pages/" + pageId);
			var response = await client.SendAsync(deleteMessage);

			return await HttpUtils.TranslateResponse(response);
		}

		#endregion
	}
}