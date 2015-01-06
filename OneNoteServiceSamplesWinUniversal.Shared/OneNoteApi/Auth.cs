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

using System;
using System.Diagnostics;
using System.Globalization;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Windows.Security.Authentication.OnlineId;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OneNoteServiceSamplesWinUniversal.OneNoteApi
{
	/// <summary>
	/// This class contains all Authentication related code for interacting with the OneNote APIs
	/// </summary>
	/// <remarks>
	/// In our previous github WinStore and WinPhone Code samples we demonstrated how to use the 
	/// LiveSDK to do OAuth against the Microsoft Account service.
	/// For this Windows 8.1 Universal code sample, we'll demonstrate an alternative way to do
	/// OAuth using the new Windows.Security.Authentication.OnlineId.OnlineIdauthenticator
	/// class.
	/// http://msdn.microsoft.com/en-us/library/windows/apps/windows.security.authentication.onlineid.onlineidauthenticator.aspx
	/// Both the existing Live SDK approach and this alternative will work in Windows 8.1 universal apps
	/// NOTE: The usage of the OnlineIdAuthenticator class is based on the Windows universal code sample at
	/// http://code.msdn.microsoft.com/windowsapps/Windows-account-authorizati-7c95e284
	/// </remarks>
	internal static class Auth
	{
		private static readonly OnlineIdAuthenticator Authenticator = new OnlineIdAuthenticator();
		private static string _accessToken;

		// TODO: Replace the below ClientId with your app's ClientId.
		// For more info, see: http://msdn.microsoft.com/en-us/library/office/dn575426(v=office.15).aspx
		private const string ClientId = "000000004011CF40";

		// OneNote APIs support multiple OAuth scopes.
		// As a guideline, always choose the least permissible 'office.onenote*' scope that your app needs.
		// Since this code sample demonstrates multiple aspects of the APIs, it uses the most
		// permissible scope: office.onenote_update.
		// TODO: Replace the below scopes with the least permissions your app needs
		private const string Scopes = "office.onenote_update wl.signin wl.offline_access";

		private const string LiveApiMeUri = "https://apis.live.net/v5.0/me?access_token=";

		/// <summary>
		/// Gets a valid authentication token. Also refreshes the access token if it has expired.
		/// </summary>
		/// <remarks>
		/// Used by the API request generators before making calls to the OneNote APIs.
		/// </remarks>
		/// <returns>valid authentication token</returns>
		internal static async Task<string> GetAuthToken()
		{
			if (String.IsNullOrWhiteSpace(_accessToken))
			{
				try
				{
					var serviceTicketRequest = new OnlineIdServiceTicketRequest(Scopes, "DELEGATION");
					var result =
						await Authenticator.AuthenticateUserAsync(new[] {serviceTicketRequest}, CredentialPromptType.PromptIfNeeded);
					if (result.Tickets[0] != null)
					{
						_accessToken = result.Tickets[0].Value;
						_accessTokenExpiration = DateTimeOffset.UtcNow.AddMinutes(AccessTokenApproxExpiresInMinutes);
					}
				}
				catch (Exception)
				{
					// Authentication failed
					if (Debugger.IsAttached) Debugger.Break();
				}
			}
			await RefreshAuthTokenIfNeeded();
			return _accessToken;
		}

		internal static async Task SignOut()
		{
			if (IsSignedIn)
			{
				_accessToken = null;
				await Authenticator.SignOutUserAsync();
			}
		}

		internal static bool IsSignedIn
		{
			get { return Authenticator != null && Authenticator.CanSignOut && _accessToken != null; }
		}

		internal static async Task<string> GetUserName()
		{
			if (_accessToken != null)
			{
				var uri = new Uri(LiveApiMeUri + _accessToken);
				var client = new HttpClient();
				var result = await client.GetAsync(uri);

				string jsonUserInfo = await result.Content.ReadAsStringAsync();
				if (jsonUserInfo != null)
				{
					var json = JObject.Parse(jsonUserInfo);
					return json["name"].ToString();
				}
			}
			return string.Empty;
		}

		#region RefreshToken related code

		// Collateral used to refresh access token (only applicable when the app uses the wl.offline_access wl.signin scopes)
		private static DateTimeOffset _accessTokenExpiration;
		private static string _refreshToken;
		private const int AccessTokenApproxExpiresInMinutes = 59;

		private const string MsaTokenRefreshUrl = "https://login.live.com/oauth20_token.srf";
		private const string TokenRefreshContentType = "application/x-www-form-urlencoded";
		private const string TokenRefreshRedirectUri = "https://login.live.com/oauth20_desktop.srf";

		private const string TokenRefreshRequestBody =
			"client_id={0}&redirect_uri={1}&grant_type=refresh_token&refresh_token={2}";

		/// <summary>
		///  Refreshes the live authentication access token if it has expired
		/// </summary>
		private static async Task RefreshAuthTokenIfNeeded()
		{
			if (_accessTokenExpiration.CompareTo(DateTimeOffset.UtcNow) <= 0)
			{
				await AttemptAccessTokenRefresh();
			}
		}

		/// <summary>
		///     This method tries to refresh the access token. The token needs to be
		///     refreshed continuously, so that the user is not prompted to sign in again
		/// </summary>
		/// <returns></returns>
		private static async Task AttemptAccessTokenRefresh()
		{
			var createMessage = new HttpRequestMessage(HttpMethod.Post, MsaTokenRefreshUrl)
			{
				Content = new StringContent(
					String.Format(CultureInfo.InvariantCulture, TokenRefreshRequestBody,
						ClientId,
						TokenRefreshRedirectUri,
						_refreshToken),
					Encoding.UTF8,
					TokenRefreshContentType)
			};

			var httpClient = new HttpClient();
			HttpResponseMessage response = await httpClient.SendAsync(createMessage);
			await ParseRefreshTokenResponse(response);
		}

		/// <summary>
		///     Handle the RefreshToken response
		/// </summary>
		/// <param name="response">The HttpResponseMessage from the TokenRefresh request</param>
		private static async Task ParseRefreshTokenResponse(HttpResponseMessage response)
		{
			if (response.StatusCode == HttpStatusCode.OK)
			{
				dynamic responseObject = JsonConvert.DeserializeObject(await response.Content.ReadAsStringAsync());
				_accessToken = responseObject.access_token;
				_accessTokenExpiration = _accessTokenExpiration.AddSeconds((double) responseObject.expires_in);
				_refreshToken = responseObject.refresh_token;
			}
		}

		#endregion
	}
}
