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
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Windows.ApplicationModel.Activation;
namespace OneNoteServiceSamplesWinUniversal.OneNoteApi
{
	/// <summary>
	/// This class contains all Authentication related code for interacting with the OneNote APIs over O365
	/// </summary>
	internal static class O365Auth
	{
		private static AuthenticationResult _authenticationResult;
		private static AuthenticationContext _authenticationContext;

		// TODO: Replace the below ClientId with your app's ClientId.
		// For more info, see: http://msdn.microsoft.com/en-us/library/office/dn575426(v=office.15).aspx
		private const string ClientId = "2bb5432a-93d1-4a3d-955e-e20c8d0e00e0"; // onebeta

		// OneNote APIs support multiple O365 scopes for OneNote entity.
		// As a guideline, always choose the least permissible scope that your app needs.
		// Since this code sample demonstrates multiple aspects of the APIs, it uses the most
		// permissible scope.

		private const string AuthContextUrl = "https://login.windows.net/Common";

		private const string ResourceUri = "https://onenote.com";

		// TODO: Replace the below RedirectUri with your app's RedirectUri.
		private const string RedirectUri = "https://localhost";

		/// <summary>
		/// Gets a valid authentication token. Also refreshes the access token if it has expired.
		/// </summary>
		/// <remarks>
		/// Used by the API request generators before making calls to the OneNote APIs.
		/// </remarks>
		/// <returns>valid authentication token</returns>
		internal static async Task<string> GetAuthToken()
		{
			return (await GetAuthenticationResult()).AccessToken;
		}

		internal static AuthenticationContext AuthContext
		{
			get
			{
				//create a new authentication context for our app
#if WINDOWS_PHONE_APP
				return _authenticationContext ?? (_authenticationContext = AuthenticationContext.CreateAsync(AuthContextUrl).GetResults());
#else
				return _authenticationContext ?? (_authenticationContext = new AuthenticationContext(AuthContextUrl));
#endif
			}
			set { _authenticationContext = value; }
		}

		internal static string AccessToken
		{
			get { return _authenticationResult != null ? _authenticationResult.AccessToken : string.Empty; }
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
			if (String.IsNullOrEmpty(AccessToken))
			{
				try
				{
					//look to see if we have an authentication context in cache already
					//we would have gotten this when we authenticated previously
					var allCachedItems = AuthContext.TokenCache.ReadItems();
					var validCachedItems = allCachedItems
									.Where(i => i.ExpiresOn > DateTimeOffset.UtcNow.UtcDateTime && IsO365Token(i.IdentityProvider))
									.OrderByDescending(e=>e.ExpiresOn);
					var cachedItem = validCachedItems.First();
					if (cachedItem != null)
					{
						//re-bind AuthenticationContext to the authority source of the cached token.
						//this is needed for the cache to work when asking for a token from that authority.
#if WINDOWS_PHONE_APP
						AuthContext = AuthenticationContext.CreateAsync(cachedItem.Authority, true).GetResults();
#else
						AuthContext = new AuthenticationContext(cachedItem.Authority, true);
#endif

						//try to get the AccessToken silently using the resourceId that was passed in
						//and the client ID of the application.
						_authenticationResult = await AuthContext.AcquireTokenSilentAsync(GetResourceHost(ResourceUri), ClientId);
						RefreshAuthTokenIfNeeded().Wait();
					}
				}
				catch (Exception)
				{
					//not in cache; we'll get it with the full oauth flow
				}
			}

			if (string.IsNullOrEmpty(AccessToken))
			{
				try
				{
					AuthContext.TokenCache.Clear();
#if WINDOWS_PHONE_APP

					_authenticationResult = await AuthContext.AcquireTokenSilentAsync(GetResourceHost(ResourceUri), ClientId);

					if (_authenticationResult == null || string.IsNullOrEmpty(_authenticationResult.AccessToken))
					{
						AuthContext.AcquireTokenAndContinue(GetResourceHost(ResourceUri), ClientId, new Uri(RedirectUri), null);
					}
#else
					_authenticationResult =
						await AuthContext.AcquireTokenAsync(GetResourceHost(ResourceUri), ClientId, new Uri(RedirectUri), PromptBehavior.Always);
#endif
				}
				catch (Exception)
				{
					// Authentication failed
					if (Debugger.IsAttached)
						Debugger.Break();
				}
			}

			return _authenticationResult;
		}

		private static bool IsO365Token(string identityProvider)
		{
			Match result =  Regex.Match(identityProvider, "https://.*.windows.*.net/.*/");
			return result.Success;
		}
		private static string GetResourceHost(string url)
		{
			Uri theHost = new Uri(url);
			return theHost.Scheme + "://" + theHost.Host + "/";
		}

		internal static async Task SignOut()
		{
			await Task.Delay(0);

			if (IsSignedIn)
			{
				if (_authenticationContext != null)
				{
					_authenticationContext.TokenCache.Clear();
				}
				_authenticationResult = null;
			}
		}

		internal static bool IsSignedIn
		{
			get { return !string.IsNullOrEmpty(AccessToken); }
		}

		internal async static Task<string> GetUserName()
		{
			await Task.Delay(0);
			return _authenticationResult != null ?
			(_authenticationResult.UserInfo != null ?
			_authenticationResult.UserInfo.GivenName  + " " + _authenticationResult.UserInfo.FamilyName: string.Empty ): string.Empty;
		}

		#region RefreshToken related code

		/// <summary>
		///  Refreshes the live authentication access token if it is about to expire in next  5 minutes
		/// </summary>
		public static async Task RefreshAuthTokenIfNeeded(double minutes = 5)
		{
			if (_authenticationResult != null && DateTimeOffset.Now.UtcDateTime.AddMinutes(minutes) > _authenticationResult.ExpiresOn)
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
			_authenticationResult = await AuthContext.AcquireTokenByRefreshTokenAsync(_authenticationResult.RefreshToken, ClientId);
		}

		#endregion

#if WINDOWS_PHONE_APP
		internal static async Task ContinueAcquireTokenAsync(WebAuthenticationBrokerContinuationEventArgs args)
		{
			_authenticationResult = await _authenticationContext.ContinueAcquireTokenAsync(args);

			// TODO by app developer: ideally we want to preserve the state of what we wanted to do before the continuation call, and do it.
			// see https://msdn.microsoft.com/library/windows/apps/dn631755.aspx and http://www.cloudidentity.com/blog/2014/06/16/adal-for-windows-phone-8-1-deep-dive/
		}
#endif
	}
}
