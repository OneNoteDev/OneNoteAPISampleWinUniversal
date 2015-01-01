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

namespace OneNoteServiceSamplesWinUniversal.OneNoteApi.Notebooks
{
	/// <summary>
	/// Class to show a selection of examples for getting notebook meta data via HTTP GET to the OneNote API
	/// - Getting a notebook info is represented via the GET HTTP verb.
	/// - Getting info for ALL notebooks the user has is represented by the Uri: https://www.onenote.com/api/v1.0/notebooks
	///    - Getting info for notebooks but with filters applied is represented by the Uri: https://www.onenote.com/api/v1.0/notebooks?$filter=...
	/// - Getting info for a specific notebook is represented by the Uri: https://www.onenote.com/api/v1.0/notebooks/{notebookId}
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
	///  client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await Auth.GetAuthToken());
	///  var getMessage = new HttpRequestMessage(HttpMethod.Get, "https://www.onenote.com/api/v1.0/notebooks");
	///  HttpResponseMessage response = await client.SendAsync(createMessage);
	/// </code>
	public static class GetNotebooksExample
	{
		#region Examples of GET https://www.onenote.com/api/v1.0/notebooks AND query params

		/// <summary>
		/// Get info for ALL  of the user's notebooks
		/// This include the user's notebooks as well as notebooks shared with this user by others.
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<List<ApiBaseResponse>> GetAllNotebooks(bool debug)
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

			// Prepare an HTTP GET request to the Notebooks endpoint
			var createMessage = new HttpRequestMessage(HttpMethod.Get, @"https://www.onenote.com/api/v1.0/notebooks");

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await TranslateListOfNotebooksResponse(response);
		}

		/// <summary>
		/// Get info for ALL notebooks matching a filter criteria
		/// In this example the filter criteria is searching for an exact matching notebook name
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="nameFilterString">search for the given notebook name</param>
		/// <returns>The converted HTTP response message</returns>
		/// <example>https://www.onenote.com/api/v1.0/notebooks?$filter=name%20eq%20'API' 
		///  returns all notebooks with name equals (case-sensitive) 'API'</example>
		public static async Task<List<ApiBaseResponse>> GetAllNotebooksWithNameMatchingFilterQueryParam(bool debug, string nameFilterString)
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

			// Prepare an HTTP GET request to the Notebooks endpoint
			var createMessage = new HttpRequestMessage(HttpMethod.Get,
				String.Format(@"https://www.onenote.com/api/v1.0/notebooks?$filter=name eq '{0}'", WebUtility.UrlEncode(nameFilterString)));

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await TranslateListOfNotebooksResponse(response);
		}

		/// <summary>
		/// Get info for ALL notebooks matching a filter criteria
		/// In this example the filter criteria is searching for the user's role as not Owner (i.e. user is a Reader/Contributor).
		/// This will return all notebooks that were shared by others with this user
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <returns>The converted HTTP response message</returns>
		/// <example>https://www.onenote.com/api/v1.0/notebooks?$filter=userRole%20ne%20Microsoft.OneNote.Api.UserRole'Owner' 
		///  returns all notebooks where userRole is not Owner. 
		/// NOTE: There is an easier way to get the same results by filtering on 'isShared' == true.
		/// </example>
		public static async Task<List<ApiBaseResponse>> GetAllNotebooksWithUserRoleAsNotOwnerFilterQueryParam(bool debug)
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

			// Prepare an HTTP GET request to the Notebooks endpoint
			var createMessage = new HttpRequestMessage(HttpMethod.Get,
				@"https://www.onenote.com/api/v1.0/notebooks?$filter=userRole ne Microsoft.OneNote.Api.UserRole'Owner'");

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await TranslateListOfNotebooksResponse(response);
		}

		/// <summary>
		/// Get selected info for ALL notebooks ordered in a given way
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="orderByFieldName">the field that is used to order the results</param>
		/// <param name="selectFieldNames">meta data fields to return in the response</param>
		/// <returns>The converted HTTP response message</returns>
		/// <example>https://www.onenote.com/api/v1.0/notebooks?$select=id,name&$orderby=createdTime
		/// returns only the id and name meta data values sorting the result chronologically by createdTime (ascending order by default)</example>
		public static async Task<List<ApiBaseResponse>> GetAllNotebooksWithOrderByAndSelectQueryParams(bool debug,
			string orderByFieldName, string selectFieldNames)
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

			// Prepare an HTTP GET request to the Notebooks endpoint
			var createMessage = new HttpRequestMessage(HttpMethod.Get,
				String.Format(@"https://www.onenote.com/api/v1.0/notebooks?$select={0}&$orderby={1}",
					WebUtility.UrlEncode(selectFieldNames),
					WebUtility.UrlEncode(orderByFieldName)));

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await TranslateListOfNotebooksResponse(response);
		}

		/// <summary>
		/// Get info for ALL notebooks and their descendants sections, sectionGroups using  $expand in one roundtrip.
		/// this is a good way to build a location picker data in one API roundtrip.
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <returns>The converted HTTP response message</returns>
		/// <example>https://www.onenote.com/api/v1.0/notebooks?$expand=sections,sectionGroups($expand=sections)
		/// </example>
		public static async Task<List<ApiBaseResponse>> GetAllNotebooksExpand(bool debug)
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

			// Prepare an HTTP GET request to the Notebooks endpoint
			var createMessage = new HttpRequestMessage(HttpMethod.Get,
				String.Format(@"https://www.onenote.com/api/v1.0/notebooks?$expand=sections,sectionGroups($expand=sections)"));

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await TranslateListOfNotebooksResponse(response);
		}

		#endregion

		#region Example of GET https://www.onenote.com/api/v1.0/notebooks/{notebookId}

		/// <summary>
		/// Get info for a specific notebook
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="notebookId">Id of the notebook for which the meta data is returned</param>
		/// <remarks>  The notebookId can be fetched from an earlier GET/POST response of Notebooks endpoint (e.g. GET https://www.onenote.com/api/v1.0/notebooks ).
		/// NOTE: Using this approach, you can still query notebooks with ALL the different params shown in examples above.
		/// </remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<ApiBaseResponse> GetASpecificNotebook(bool debug, string notebookId)
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

			// Prepare an HTTP GET request to the Notebooks endpoint
			var createMessage = new HttpRequestMessage(HttpMethod.Get, @"https://www.onenote.com/api/v1.0/notebooks/" + notebookId);

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await TranslateNotebookResponse(response);
		}

		#endregion

		#region Helper methods used in the examples

		/// <summary>
		/// Convert the HTTP response message into a simple structure suitable for apps to process
		/// </summary>
		/// <param name="response">The response to convert</param>
		/// <returns>A simple response</returns>
		private static async Task<List<ApiBaseResponse>> TranslateListOfNotebooksResponse(HttpResponseMessage response)
		{
			var apiBaseResponse = new List<ApiBaseResponse>();
			string body = await response.Content.ReadAsStringAsync();
			if (response.StatusCode == HttpStatusCode.OK
				/* GET Notebooks calls always return 200-OK upon success */)
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
		private static async Task<ApiBaseResponse> TranslateNotebookResponse(HttpResponseMessage response)
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
