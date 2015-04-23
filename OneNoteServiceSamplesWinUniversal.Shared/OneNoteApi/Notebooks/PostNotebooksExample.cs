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

namespace OneNoteServiceSamplesWinUniversal.OneNoteApi.Notebooks
{
	/// <summary>
	/// Class to show a selection of examples creating notebooks via HTTP POST to the OneNote API
	/// - Creating a new notebook is represented via the POST HTTP verb.
	/// - Creating a new notebook represented by the Uri: https://www.onenote.com/api/v1.0/notebooks
	/// For more info, see http://dev.onenote.com/docs
	/// </summary>
	/// <remarks>
	/// NOTE: All notebooks get created under the 'Documents' folder of the user's OneDrive account.
	/// The notebook name is specified in the request body
	/// NOTE: It is not the goal of this code sample to produce well re-factored, elegant code.
	/// You may notice code blocks being duplicated in various places in this project.
	/// We have deliberately added these code blocks to allow anyone browsing the sample
	/// to easily view all related functionality in near proximity.
	/// </remarks>
	/// <code>
	///  var client = new HttpClient();
	///  client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
	///  client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await Auth.GetAuthToken());
	///  var createMessage = new HttpRequestMessage(HttpMethod.Post, "https://www.onenote.com/api/v1.0/notebooks")
	///  {
	///      Content = new StringContent("{name: NewNotebookName }", System.Text.Encoding.UTF8, "application/json")
	///  };
	///  HttpResponseMessage response = await client.SendAsync(createMessage);
	/// </code>
	public static class PostNotebooksExample
	{
		#region Examples of POST https://www.onenote.com/api/v1.0/notebooks

		/// <summary>
		/// Create a notebook with a given name
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="notebookName"></param>
		/// <param name="provider"></param>
		/// <param name="apiRoute"></param>
		/// <remarks>Create notebook using a application/json content type</remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<ApiBaseResponse> CreateSimpleNotebook(bool debug, string notebookName, AuthProvider provider, string apiRoute)
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

			// Prepare an HTTP POST request to the Notebooks endpoint
			// The request body content type is application/json and require a name property
			var createMessage = new HttpRequestMessage(HttpMethod.Post, apiRoute + "notebooks")
			{
				Content = new StringContent("{ name : '" + WebUtility.UrlEncode(notebookName) + "' }", Encoding.UTF8, "application/json")
			};

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await HttpUtils.TranslateResponse(response);
		}

		#endregion

	}
}
