using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace OneNoteServiceSamplesWinUniversal.OneNoteApi
{
    public static class HttpUtils
    {
		#region Helper methods used in the examples

	    /// <summary>
	    /// Convert the HTTP response message into a simple structure suitable for apps to process
	    /// </summary>
	    /// <param name="response">The response to convert</param>
	    /// <param name="expectedStatusCode"></param>
	    /// <returns>A simple response</returns>
	    public static async Task<ApiBaseResponse> TranslateResponse(HttpResponseMessage response, HttpStatusCode expectedStatusCode = HttpStatusCode.Created)
	    {
		    ApiBaseResponse apiBaseResponse;
		    string body = await response.Content.ReadAsStringAsync();
			if (response.StatusCode == expectedStatusCode
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
