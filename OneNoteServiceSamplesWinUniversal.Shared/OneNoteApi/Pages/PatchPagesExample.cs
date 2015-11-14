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

using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace OneNoteServiceSamplesWinUniversal.OneNoteApi.Pages
{
	public class PatchCommand
	{
		//The target to update
		public string Target { get; set; }

		//The action to perform which can be insert, append or replace
		public string Action { get; set; }

		//The new HTML content to be added to the page for the action
		public string Content { get; set; }
	}

	/// <summary>
	/// Class to show a selection of examples updating pages via HTTP PATCH requests to the OneNote API
	/// Updating an existing page is achieved by sending HTTP PATCH requests to the URL: https://www.onenote.com/api/v1.0/pages/{id}/content
	/// For more info, see http://dev.onenote.com/docs
	/// </summary>
	/// <remarks>
	/// NOTE: It is not the goal of this code sample to produce well re-factored, elegant code.
	/// You may notice code blocks being duplicated in various places in this project.
	/// We have deliberately added these code blocks to allow anyone browsing the sample
	/// to easily view all related functionality in near proximity.
	/// </remarks>
	/// <code>
	///  var client = new HttpClient();
	///  client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
	///  client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await Auth.GetAuthToken());
	/// 
	///  var createMessage = new HttpRequestMessage(new HttpMethod("PATCH"), "https://www.onenote.com/api/v1.0/pages/{pageId}/content")
	///  {
	///      Content = new StringContent("[{ 'target': 'body', 'action': 'append', 'content': '<p>New trailing content</p>' }]", Encoding.UTF8, "application/json")
	///  };
	///  HttpResponseMessage response = await client.SendAsync(createMessage);
	/// </code>
	public static class PatchPagesExample
	{
		#region Example of PATCH https://www.onenote.com/api/v1.0/pages/{id}/content

		/// <summary>
		/// Append to the default outline in the content of an existing page
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="pageId">The identifier of the page whose content to append to</param>
		/// <param name="provider"></param>
		/// <param name="apiRoute"></param>
		/// <remarks>Create page using a single part text/html content type</remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<ApiBaseResponse> AppendToDefaultOutlineInPageContent(bool debug, string pageId, AuthProvider provider, string apiRoute)
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

			//The request content must be a JSON list of commands
			var commands = new List<PatchCommand>();

			//Each command must contain atleast a target and an action. Specifying content may be optional or mandatory depending on the action specified.
			var appendCommand = new PatchCommand();
			appendCommand.Target = "body";
			//While in this case the PATCH action is append, the API supports replace (to replace existing page content) and insert (insert new page content relative to existing content) actions as well.
			appendCommand.Action = "append";
			appendCommand.Content = "<p>New trailing content</p>";

			commands.Add(appendCommand);

			var appendRequestContent = JsonConvert.SerializeObject(commands);

			// Prepare an HTTP PATCH request to the Pages/{pageId}/content endpoint
			// The request body content type is application/json
			var createMessage = new HttpRequestMessage(new HttpMethod("PATCH"), apiRoute + "pages/" + pageId + "/content")
			{
				Content = new StringContent(appendRequestContent, Encoding.UTF8, "application/json")
			};

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await TranslatePatchResponse(response);
		}

		#endregion

		#region Helper methods used in the examples

		/// <summary>
		/// Convert the HTTP response message into a simple structure suitable for apps to process
		/// </summary>
		/// <param name="response">The response to convert</param>
		/// <returns>A simple response</returns>
		private static async Task<ApiBaseResponse> TranslatePatchResponse(HttpResponseMessage response)
		{
			var apiBaseResponse = new ApiBaseResponse();;
			string body = await response.Content.ReadAsStringAsync();

			// PATCH returns 204 NoContent upon success
			apiBaseResponse.StatusCode = response.StatusCode;

			//If the request was successful, the response body will be empty.
			//Otherwise, it will contain errors or warnings relevant to your request
			apiBaseResponse.Body = body;

			// Extract the correlation id.  Apps should log this if they want to collect the data to diagnose failures with Microsoft support 
			IEnumerable<string> correlationValues;
			if (response.Headers.TryGetValues("X-CorrelationId", out correlationValues))
			{
				apiBaseResponse.CorrelationId = correlationValues.FirstOrDefault();
			}
			
			return apiBaseResponse;
		}

		#endregion
	}
}
