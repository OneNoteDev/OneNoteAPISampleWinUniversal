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
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace OneNoteServiceSamplesWinUniversal.OneNoteApi.Sections
{
	/// <summary>
	/// Class to show a selection of examples for getting section meta data via HTTP GET to the OneNote API
	/// - Getting a section info is represented via the GET HTTP verb.
	/// - Getting info for ALL sections the user has is represented by the Uri: https://www.onenote.com/api/v1.0/sections
	///    - Getting info for sections but with filters applied is represented by the Uri: https://www.onenote.com/api/v1.0/sections?$filter=...
	/// - Getting info for ALL sections under a specific notebook is represented by the Uri: https://www.onenote.com/api/v1.0/notebooks/{notebookId}/sections
	/// - Getting info for ALL sections under a specific sectionGroup is represented by the Uri: https://www.onenote.com/api/v1.0/sectionGroups/{sectionGroupId}/sections
	/// - Getting info for a specific section is represented by the Uri: https://www.onenote.com/api/v1.0/sections/{sectionId}
	/// For more info, see http://dev.onenote.com/docs
	/// </summary>
	/// <remarks>
	/// NOTE: Our APIs support OData-style querying so you can use query params like $filter, $orderby, $top, $skip, $select etc.
	/// For more info, see http://www.odata.org/getting-started/basic-tutorial/
	/// NOTE: It is not the goal of this code sample to produce well re-factored, elegant code.
	/// You may notice code blocks being duplicated in various places in this project.
	/// We have deliberately added these code blocks to allow anyone browsing the sample
	/// to easily view all related functionality in near proximity.
	/// </remarks>
	/// <code>
	///  var client = new HttpClient();
	///  client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
	///  client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await Auth.GetAuthprovider());
	///  var getMessage = new HttpRequestMessage(HttpMethod.Get, "https://www.onenote.com/api/v1.0/sections");
	///  HttpResponseMessage response = await client.SendAsync(createMessage);
	/// </code>
	public static class GetSectionsExample
	{
		#region Examples of GET https://www.onenote.com/api/v1.0/sections AND query params

		/// <summary>
		/// Get info for ALL  of the user's sections
		/// This include the user's sections as well as sections shared with this user by others.
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="provider"></param>
		/// <param name="apiRoute"></param>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<List<ApiBaseResponse>> GetAllSections(bool debug, AuthProvider provider, string apiRoute)
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

			// Prepare an HTTP GET request to the Sections endpoint
			var createMessage = new HttpRequestMessage(HttpMethod.Get, apiRoute + "sections");

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await TranslateListOfSectionsResponse(response);
		}

		/// <summary>
		/// Get info for ALL sections matching a filter criteria
		/// In this example the filter criteria is searching for an exact matching section name
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="nameFilterString">search for the given section name</param>
		/// <param name="provider"></param>
		/// <param name="apiRoute"></param>
		/// <returns>The converted HTTP response message</returns>
		/// <example>https://www.onenote.com/api/v1.0/sections?$filter=name%20eq%20'API' 
		///  returns all sections with name equals (case-sensitive) 'API'</example>
		public static async Task<List<ApiBaseResponse>> GetAllSectionsWithNameMatchingFilterQueryParam(bool debug, string nameFilterString, AuthProvider provider, string apiRoute)
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

			// Prepare an HTTP GET request to the Sections endpoint
			var createMessage = new HttpRequestMessage(HttpMethod.Get,
				String.Format(apiRoute + "sections?$filter=name eq '{0}'", WebUtility.UrlEncode(nameFilterString)));

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await TranslateListOfSectionsResponse(response);
		}

		#endregion

		#region Example of GET https://www.onenote.com/api/v1.0/notebooks/{notebookId}/sections

		/// <summary>
		/// Get info for all sections under a given notebook Id
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="notebookId">Id of the parent notebook</param>
		/// <param name="provider"></param>
		/// <param name="apiRoute"></param>
		/// <remarks>  The notebookId can be fetched from an earlier GET/POST response of Notebooks endpoint (e.g. GET https://www.onenote.com/api/v1.0/notebooks ).
		/// NOTE: Using this approach, you can still query sections with ALL the different params shown in examples above.
		/// </remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<List<ApiBaseResponse>> GetSectionsUnderASpecificNotebook(bool debug, string notebookId, AuthProvider provider, string apiRoute)
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

			// Prepare an HTTP GET request to the Sections endpoint
			var createMessage = new HttpRequestMessage(HttpMethod.Get, apiRoute + "notebooks/" + notebookId + "/sections");

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await TranslateListOfSectionsResponse(response);
		}

		#endregion

		#region Example of GET https://www.onenote.com/api/v1.0/sectionGroups/{sectionGroupId}/sections

		/// <summary>
		/// Get info for sections under a specific sectionGroup
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="sectionGroupId">Id of the sectionGroup for which the meta data is returned</param>
		/// <param name="apiRoute"></param>
		/// <remarks>  The sectionGroupId can be fetched from an earlier GET/POST response of SectionGroups endpoint (e.g. GET https://www.onenote.com/api/v1.0/sectiongroups ).
		/// NOTE: Using this approach, you can still query sections with ALL the different params shown in examples above.
		/// </remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<List<ApiBaseResponse>> GetSectionsUnderASpecificSectionGroup(bool debug, string sectionGroupId, string apiRoute)
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
				await Auth.GetAuthToken());

			// Prepare an HTTP GET request to the Sections endpoint
			var createMessage = new HttpRequestMessage(HttpMethod.Get, apiRoute + "sectionGroups/" + sectionGroupId + "/sections");

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await TranslateListOfSectionsResponse(response);
		}

		#endregion

		#region Example of GET https://www.onenote.com/api/v1.0/sections/{sectionId}

		/// <summary>
		/// Get info for a specific section
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="sectionId">Id of the section for which the meta data is returned</param>
		/// <param name="provider"></param>
		/// <param name="apiRoute"></param>
		/// <remarks>  The sectionId can be fetched from an earlier GET/POST response of Sections/Notebooks/SectionGroups endpoint (e.g. GET https://www.onenote.com/api/v1.0/sections ).
		/// NOTE: Using this approach, you can still query sections with ALL the different params shown in examples above.
		/// </remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<ApiBaseResponse> GetASpecificSection(bool debug, string sectionId, AuthProvider provider, string apiRoute)
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

			// Prepare an HTTP GET request to the Sections endpoint
			var createMessage = new HttpRequestMessage(HttpMethod.Get, apiRoute + "sections/" + sectionId);

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await TranslateSectionsResponse(response);
		}

		#endregion

		#region Helper methods used in the examples

		/// <summary>
		/// Convert the HTTP response message into a simple structure suitable for apps to process
		/// </summary>
		/// <param name="response">The response to convert</param>
		/// <returns>A simple response</returns>
		private static async Task<List<ApiBaseResponse>> TranslateListOfSectionsResponse(HttpResponseMessage response)
		{
			var apiBaseResponse = new List<ApiBaseResponse>();
			string body = await response.Content.ReadAsStringAsync();
			if (response.StatusCode == HttpStatusCode.OK
				/* GET Sections calls always return 200-OK upon success */)
			{
					var content = JObject.Parse(body);
					apiBaseResponse = new List<ApiBaseResponse>(JsonConvert.DeserializeObject<List<GenericEntityResponse>>(content["value"].ToString()));
			}
			if (apiBaseResponse.Count == 0)
			{
				apiBaseResponse.Add(new ApiBaseResponse());
			}

			// Extract the correlation id.  Apps should log this if they want to collect the data to diagnose failures with Microsoft support 
			IEnumerable<string> correlationValues;
			if (response.Headers.TryGetValues("X-CorrelationId", out correlationValues))
			{
				apiBaseResponse[0].CorrelationId = correlationValues.FirstOrDefault();
			}
			apiBaseResponse[0].StatusCode = response.StatusCode;
			apiBaseResponse[0].Body = body;
			return apiBaseResponse;
		}
		/// <summary>
		/// Convert the HTTP response message into a simple structure suitable for apps to process
		/// </summary>
		/// <param name="response">The response to convert</param>
		/// <returns>A simple response</returns>
		private static async Task<ApiBaseResponse> TranslateSectionsResponse(HttpResponseMessage response)
		{
			ApiBaseResponse apiBaseResponse;
			string body = await response.Content.ReadAsStringAsync();
			if (response.StatusCode == HttpStatusCode.OK
				/* GET Pages calls always return 200-OK upon success */)
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
