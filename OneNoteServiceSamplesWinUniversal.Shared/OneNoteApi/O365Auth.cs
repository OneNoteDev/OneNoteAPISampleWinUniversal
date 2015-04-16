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
using System.Linq;
using System.Threading.Tasks;

using Microsoft.IdentityModel.Clients.ActiveDirectory;
namespace OneNoteServiceSamplesWinUniversal.OneNoteApi
{
	/// <summary>
	/// This class contains all Authentication related code for interacting with the OneNote APIs over O365
	/// </summary>
	internal static class O365Auth
	{
		private static AuthenticationResult _authenticationResult;
		private static AuthenticationContext _authenticationContext;
		private static string _accessToken;

		// TODO: Replace the below ClientId with your app's ClientId.
		// For more info, see: http://msdn.microsoft.com/en-us/library/office/dn575426(v=office.15).aspx
		private const string ClientIdPpe = "d1b48acb-5ff2-4e43-b6a2-96e4e5ab7471"; // Gareth's tenant
		private const string ClientIdProd = "172c4773-9607-480b-b94a-29c48d998080"; // Sharad's tenant

		private static string ClientId 
		{
			get
			{
				return !ProductionReady ? ClientIdPpe : ClientIdProd; 
			}
		}

	// OneNote APIs support multiple O365 scopes for OneNote entity.
		// As a guideline, always choose the least permissible scope that your app needs.
		// Since this code sample demonstrates multiple aspects of the APIs, it uses the most
		// permissible scope.

		private const string AuthContextUrlProduction = "https://login.windows.net/Common";
		private const string AuthContextUrlPpe = "https://login.windows-ppe.net/Common";

		private static string AuthContextUrl
		{
			get
			{
				return !ProductionReady ? AuthContextUrlPpe : AuthContextUrlProduction;
			}
		}

		private const string ResourceUriProd = "https://sharmas.sharepoint.com";
		private const string ResourceUriPpe = "https://onenote.com";
		private static string ResourceUri
		{
			get
			{
				return GetResourceHost(!ProductionReady ? ResourceUriPpe : ResourceUriProd);
			}
		}

		private const string RedirectUriProd = "http://TodoListClient";
		private const string RedirectUriPpe = "https://localhost";

		private static string RedirectUri
		{
			get
			{
				return !ProductionReady ? RedirectUriPpe : RedirectUriProd;
			}
		}


		internal static bool ProductionReady
		{
			get { return _productionReady; }
			set { _productionReady = value; }
		}

		/// <summary>
		/// Gets a valid authentication token. Also refreshes the access token if it has expired.
		/// </summary>
		/// <remarks>
		/// Used by the API request generators before making calls to the OneNote APIs.
		/// </remarks>
		/// <returns>valid authentication token</returns>
		internal static async Task<string> GetAuthToken()
		{
			_authenticationResult = await GetAuthenticationResult();
			return _accessToken;
		}

		internal static AuthenticationContext AuthenticationContext
		{
			get
			{
				//create a new authentication context for our app
				return _authenticationContext ?? (_authenticationContext = new AuthenticationContext(AuthContextUrl));
			}
			set { _authenticationContext = value; }
		}
		/// <summary>
		/// Gets a valid authentication token. Also refreshes the access token if it has expired.
		/// </summary>
		/// <remarks>
		/// Used by the API request generators before making calls to the OneNote APIs.
		/// </remarks>
		/// <returns>valid authentication token</returns>
		internal static async Task<AuthenticationResult> GetAuthenticationResult()
		{
			AuthenticationResult authenticationResult = null;


			if (String.IsNullOrWhiteSpace(_accessToken))
			{
				try
				{
					//look to see if we have an authentication context in cache already
					//we would have gotten this when we authenticated previously
					if (AuthenticationContext.TokenCache.ReadItems().Any(i => i.ExpiresOn > DateTimeOffset.UtcNow.UtcDateTime))
					{
						//re-bind AuthenticationContext to the authority source of the cached token.
						//this is needed for the cache to work when asking for a token from that authority.
						string cachedAuthority =
							AuthenticationContext.TokenCache.ReadItems().First().Authority;
						//create a new authentication context for our app
						AuthenticationContext = new AuthenticationContext(cachedAuthority);

						//try to get the AccessToken silently using the resourceId that was passed in
						//and the client ID of the application.
						authenticationResult = await AuthenticationContext.AcquireTokenSilentAsync(ResourceUri, ClientId);

						_accessToken = authenticationResult.AccessToken;
						_refreshToken = authenticationResult.RefreshToken;
					}
				}
				catch (Exception)
				{
					//not in cache; we'll get it with the full oauth flow
				}

				if (authenticationResult == null || string.IsNullOrEmpty(authenticationResult.AccessToken))
				{
					try
					{
						await SignOut();
						authenticationResult = await AuthenticationContext.AcquireTokenAsync(ResourceUri, ClientId, new Uri(RedirectUri), PromptBehavior.Auto);

						_accessToken = authenticationResult.AccessToken;
						_refreshToken = authenticationResult.RefreshToken;
					}
					catch (Exception)
					{
						// Authentication failed
						if (Debugger.IsAttached)
							Debugger.Break();
					}
				}
				else
				{
					RefreshAuthTokenIfNeeded().Wait();
				}

			}

			return authenticationResult;
		}

		private static string GetResourceHost(string url)
		{
			Uri theHost = new Uri(url);
			return theHost.Scheme + "://" + theHost.Host + "/";
		}

#pragma warning disable 1998
		internal static async Task SignOut()
#pragma warning restore 1998
		{
			if (IsSignedIn && _authenticationResult != null && _authenticationResult.UserInfo != null && !string.IsNullOrEmpty(_authenticationResult.UserInfo.UniqueId))
			{
				_authenticationContext.TokenCache.Clear();
				_accessToken = null;

			}
		}

		internal static bool IsSignedIn
		{
			get { return _authenticationResult != null && _accessToken != null; }
		}

		internal async static Task<string> GetUserName()
		{
			return _authenticationResult != null ?(_authenticationResult.UserInfo != null ?_authenticationResult.UserInfo.GivenName: string.Empty ): string.Empty;
		}

		#region RefreshToken related code

		// Collateral used to refresh access token
		private static DateTimeOffset _accessTokenExpiration;
		private static string _refreshToken;
		private static bool _productionReady;// TODO: - when ready for production, set this to true by default

		/// <summary>
		///  Refreshes the live authentication access token if it is about to expire in next  5 minutes
		/// </summary>
		public static async Task RefreshAuthTokenIfNeeded(double minutes = 5)
		{
			if (DateTimeOffset.Now.UtcDateTime.AddMinutes(minutes) > _accessTokenExpiration)
			{
				await AttemptAccessTokenRefresh();
			}
		}

		/// <summary>
		/// This method tries to refresh the access token. The token needs to be
		/// efreshed continuously, so that the user is not prompted to sign in again
		/// </summary>
		/// <returns></returns>
		public static async Task AttemptAccessTokenRefresh()
		{
			_authenticationResult = await AuthenticationContext.AcquireTokenByRefreshTokenAsync(_refreshToken, ClientId);
			_accessToken = _authenticationResult.AccessToken;
			_accessTokenExpiration = _authenticationResult.ExpiresOn;
		}

		#endregion
	}
}
