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

using Newtonsoft.Json;
using System.Collections.Generic;
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
	/// Class to show a selection of example creating sections via HTTP POST to the OneNote API
	/// - Creating a new section is represented via the POST HTTP verb.
	/// - Creating a new section under a given notebook is represented by the Uri: https://www.onenote.com/api/v1.0/notebooks/{notebookId}/sections
	/// For more info, see http://dev.onenote.com/docs
	/// </summary>
	/// <remarks>
	/// NOTE: All sections require a parent notebook
	/// The section name is specified in the request body
	/// NOTE: It is not the goal of this code sample to produce well re-factored, elegant code.
	/// You may notice code blocks being duplicated in various places in this project.
	/// We have deliberately added these code blocks to allow anyone browsing the sample
	/// to easily view all related functionality in near proximity.
	/// </remarks>
	/// <code>
	///  var client = new HttpClient();
	///  client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
	///  client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await Auth.GetAuthToken());
	///  var createMessage = new HttpRequestMessage(HttpMethod.Post, "https://www.onenote.com/api/v1.0/notebooks/{notebookId}/ections")
	///  {
	///      Content = new StringContent("{name: NewSectionName }", System.Text.Encoding.UTF8, "application/json")
	///  };
	///  HttpResponseMessage response = await client.SendAsync(createMessage);
	/// </code>
	public static class PostSectionsExample
	{
		#region Examples of POST https://www.onenote.com/api/v1.0/notebooks/{notebookId}/sections

		/// <summary>
		/// Create a section with a given name under a given notebookId
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="notebookId">parent notebook's Id</param>
		/// <param name="sectionName">name of the section to create</param>
		/// <param name="provider"></param>
		/// <param name="apiRoute"></param>
		/// <remarks>Create section using a application/json content type</remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<ApiBaseResponse> CreateSimpleSection(bool debug, string notebookId, string sectionName, AuthProvider provider, string apiRoute)
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

			// Prepare an HTTP POST request to the Sections endpoint
			// The request body content type is application/json and require a name property
			var createMessage = new HttpRequestMessage(HttpMethod.Post, @"https://www.onenote.com/api/v1.0/notebooks/" + notebookId + "/sections")
			{
				Content = new StringContent("{ name : '" + WebUtility.UrlEncode(sectionName) + "' }", Encoding.UTF8, "application/json")
			};

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await TranslateResponse(response);
		}

		#endregion

		#region Helper methods used in the examples

		/// <summary>
		/// Convert the HTTP response message into a simple structure suitable for apps to process
		/// </summary>
		/// <param name="response">The response to convert</param>
		/// <returns>A simple response</returns>
		private static async Task<ApiBaseResponse> TranslateResponse(HttpResponseMessage response)
		{
			ApiBaseResponse apiBaseResponse;
			string body = await response.Content.ReadAsStringAsync();
			if (response.StatusCode == HttpStatusCode.Created
				/* POST calls always return 201-Created upon success */)
			{
				apiBaseResponse = JsonConvert.DeserializeObject<GenericEntityResponse>(body);
			}
			else
			{
				apiBaseResponse = new ApiBaseResponse();
			}

			// Extract the correlation id.  Apps should log this if they want to collect the data to diagnose failures with Microsoft support 
			IEnumerable<string> correlationValues;
			if (response.Headers.TryGetValues("X-CorrelationId", out correlationValues))
			{
				apiBaseResponse.CorrelationId = correlationValues.FirstOrDefault();
			}
			apiBaseResponse.StatusCode = response.StatusCode;
			apiBaseResponse.Body = body;
			return apiBaseResponse;
		}

		#endregion
	}
}
