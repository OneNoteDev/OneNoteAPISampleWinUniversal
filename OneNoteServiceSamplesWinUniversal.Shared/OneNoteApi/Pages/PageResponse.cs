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

namespace OneNoteServiceSamplesWinUniversal.OneNoteApi.Pages
{
	/// <summary>
	/// This class represents the OneNote API Pages response
	/// Any response from the Pages API (POST/GET etc) can be translated into this object for ease of use.
	/// </summary>
	/// <remarks>
	/// This is not meant to be a comprehensive SDK or data model.
	/// This is ONLY a light-weight representation of a OneNote API's page response.
	/// The API's HTTP json response is deserialized into this object
	/// </remarks>
	public class PageResponse : ApiBaseResponse
	{
		/// <summary>
		/// Title of the page
		/// </summary>
		public string Title;

		public override string ToString()
		{
			return "Title: " + Title + ", Id: " + Id;
		}
	}
}
