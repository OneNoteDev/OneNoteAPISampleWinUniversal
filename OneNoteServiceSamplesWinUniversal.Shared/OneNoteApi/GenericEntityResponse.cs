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

namespace OneNoteServiceSamplesWinUniversal.OneNoteApi
{
	/// <summary>
	/// This class represents a generic the OneNote API entity response
	/// Any response from the Notebooks/Sections/SectionGroups API (POST/GET etc) can be translated into this object for ease of use.
	/// </summary>
	/// <remarks>
	/// This is not meant to be a comprehensive SDK or data model.
	/// This is ONLY a light-weight representation of a OneNote API's entities (Notebooks, Sections, SectionGroups)
	/// The API's HTTP json response is deserialized into this object
	/// </remarks>
	public class GenericEntityResponse : ApiBaseResponse
	{
		/// <summary>
		/// Name of the entity
		/// </summary>
		public string Name;

		/// <summary>
		/// Self link to the given entity
		/// </summary>
		public string Self { get; set; }

		public List<GenericEntityResponse> Sections { get; set; }

		public List<GenericEntityResponse> SectionGroups { get; set; }

		public override string ToString()
		{
			return "Name: " + Name + ", Id: " + Id;
		}
	}
}
