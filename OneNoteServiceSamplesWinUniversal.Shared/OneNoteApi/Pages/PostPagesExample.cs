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
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Windows.ApplicationModel;
using Newtonsoft.Json;

namespace OneNoteServiceSamplesWinUniversal.OneNoteApi.Pages
{
	/// <summary>
	/// Class to show a selection of examples creating pages via HTTP POST to the OneNote API
	/// - Creating a new page is represented via the POST HTTP verb.
	/// - Creating a page in the default location is represented by the Uri: https://www.onenote.com/api/v1.0/pages
	///    - Creating a page in the default notebook but under a specific named section is represented by the Uri: https://www.onenote.com/api/v1.0/pages?sectionName={SectionName}
	/// - Creating a page in a specific section is represented by the Uri: https://www.onenote.com/api/v1.0/sections/{sectionId}/pages
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
	///  string simpleHtml = "<html><head><title>Page Title</title></head>" +
	///                      "<body><p>Hello World!</p></body></html>";
	///  var createMessage = new HttpRequestMessage(HttpMethod.Post, "https://www.onenote.com/api/v1.0/pages")
	///  {
	///      Content = new StringContent(simpleHtml, System.Text.Encoding.UTF8, "text/HTML")
	///  };
	///  HttpResponseMessage response = await client.SendAsync(createMessage);
	/// </code>
	public static class PostPagesExample
	{
		#region Examples of POST https://www.onenote.com/api/v1.0/pages with different contents

		/// <summary>
		/// Create a very simple page with some formatted text.
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="sectionId"></param>
		/// <param name="provider"></param>
		/// <param name="apiRoute"></param>
		/// <remarks>Create page using a single part text/html content type</remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<ApiBaseResponse> CreateSimplePage(bool debug, AuthProvider provider, string apiRoute)
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
			client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await Auth.GetAuthToken(provider));

			string date = GetDate();
			string simpleHtml = "<html>" +
								"<head>" +
								"<title>A simple page created from basic HTML-formatted text</title>" +
								"<meta name=\"created\" content=\"" + date + "\" />" +
								"</head>" +
								"<body>" +
								"<p>This is a page that just contains some simple <i>formatted</i> <b>text</b></p>" +
								"<p>Here is a <a href=\"http://www.microsoft.com\">link</a></p>" +
								"</body>" +
								"</html>";

			// Prepare an HTTP POST request to the Pages endpoint
			// The request body content type is text/html
			var createMessage = new HttpRequestMessage(HttpMethod.Post, apiRoute + "/pages")
			{
				Content = new StringContent(simpleHtml, Encoding.UTF8, "text/html")
			};

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await HttpUtils.TranslateResponse(response);
		}

		/// <summary>
		/// Create a page with an image on it.
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="sectionId"></param>
		/// <param name="provider"></param>
		/// <param name="apiRoute"></param>
		/// <remarks>Create page using a multipart/form-data content type</remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<ApiBaseResponse> CreatePageWithImage(bool debug, string sectionId, AuthProvider provider, string apiRoute)
		{
			if (debug)
			{
				Debugger.Launch();
				Debugger.Break();
			}

			var client = new HttpClient();

			// Note: API only supports JSON return type.
			client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

			// Not adding the Authentication header would produce an unauthorized call and the API will return a 401
			client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer",
				await Auth.GetAuthToken(provider));

			const string imagePartName = "image1";
			string date = GetDate();
			string simpleHtml = "<html>" +
								"<head>" +
								"<title>A simple page created with an image on it</title>" +
								"<meta name=\"created\" content=\"" + date + "\" />" +
								"</head>" +
								"<body>" +
								"<h1>This is a page with an image on it</h1>" +
								"<img src=\"name:" + imagePartName +
								"\" alt=\"A beautiful logo\" width=\"426\" height=\"68\" />" +
								"</body>" +
								"</html>";

			// Create the image part - make sure it is disposed after we've sent the message in order to close the stream.
			HttpResponseMessage response;
			using (var imageContent = new StreamContent(await GetBinaryStream("assets\\Logo.jpg")))
			{
				imageContent.Headers.ContentType = new MediaTypeHeaderValue("image/jpeg");
				var createMessage = new HttpRequestMessage(HttpMethod.Post, apiRoute + "sections/" + sectionId + "/pages")
				{
					Content = new MultipartFormDataContent
					{
						{new StringContent(simpleHtml, Encoding.UTF8, "text/html"), "Presentation"},
						{imageContent, imagePartName}
					}
				};

				// Must send the request within the using block, or the image stream will have been disposed.
				response = await client.SendAsync(createMessage);
			}

			return await HttpUtils.TranslateResponse(response);
		}

		/// <summary>
		/// Create a page with a file attachment
		/// </summary>
		/// <param name="debug">Determines whether to execute this method under the debugger</param>
		/// <param name="sectionId"></param>
		/// <param name="provider"></param>
		/// <param name="apiRoute"></param>
		/// <remarks>Create page using a multipart/form-data content type</remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<ApiBaseResponse> CreatePageWithAttachedFile(bool debug, string sectionId, AuthProvider provider, string apiRoute)
		{
			if (debug)
			{
				Debugger.Launch();
				Debugger.Break();
			}

			var client = new HttpClient();
			string date = GetDate();
			//Note: API only supports JSON return type
			client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

			// Not adding the Authentication header would produce an unauthorized call and the API will return a 401
			client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer",
				await Auth.GetAuthToken(provider));

			const string attachmentPartName = "pdfattachment1";
			string attachmentRequestHtml = "<html>" +
										   "<head>" +
										   "<title>A page created with a file attachment</title>" +
										   "<meta name=\"created\" content=\"" + date + "\" />" +
										   "</head>" +
										   "<body>" +
										   "<h1>This is a page with a pdf file attachment</h1>" +
										   "<img data-render-src=\"name:" + attachmentPartName + "\" />" +
										   "<br />" +
										   "<object data-attachment=\"attachment.pdf\" data=\"name:" +
										   attachmentPartName + "\" />" +
										   "</body>" +
										   "</html>";

			HttpResponseMessage response;
			using (var attachmentContent = new StreamContent(await GetBinaryStream("assets\\attachment.pdf")))
			{
				attachmentContent.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
				var createMessage = new HttpRequestMessage(HttpMethod.Post, apiRoute + "sections/" + sectionId + "/pages")
				{
					Content = new MultipartFormDataContent
					{
						{new StringContent(attachmentRequestHtml, Encoding.UTF8, "text/html"), "Presentation"},
						{attachmentContent, attachmentPartName}
					}
				};
				// Must send the request within the using block, or the binary stream will have been disposed.
				response = await client.SendAsync(createMessage);
			}
			return await HttpUtils.TranslateResponse(response);
		}

		/// <summary>
		/// Create a page with an image of an embedded web page on it.
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="sectionId"></param>
		/// <param name="provider"></param>
		/// <param name="apiRoute"></param>
		/// <remarks>Create page using a multipart/form-data content type</remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<ApiBaseResponse> CreatePageWithEmbeddedWebPage(bool debug, string sectionId, AuthProvider provider, string apiRoute)
		{
			if (debug)
			{
				Debugger.Launch();
				Debugger.Break();
			}

			var client = new HttpClient();

			// Note: API only supports JSON return type.
			client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

			// Not adding the Authentication header would produce an unauthorized call and the API will return a 401
			client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer",
				await Auth.GetAuthToken(provider));

			const string embeddedPartName = "embedded1";
			const string embeddedWebPage =
				"<html>" +
				"<head>" +
				"<title>An embedded webpage</title>" +
				"</head>" +
				"<body>" +
				"<h1>This is a screen grab of a web page</h1>" +
				"<p>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam vehicula magna quis mauris accumsan, nec imperdiet nisi tempus. Suspendisse potenti. " +
				"Duis vel nulla sit amet turpis venenatis elementum. Cras laoreet quis nisi et sagittis. Donec euismod at tortor ut porta. Duis libero urna, viverra id " +
				"aliquam in, ornare sed orci. Pellentesque condimentum gravida felis, sed pulvinar erat suscipit sit amet. Nulla id felis quis sem blandit dapibus. Ut " +
				"viverra auctor nisi ac egestas. Quisque ac neque nec velit fringilla sagittis porttitor sit amet quam.</p>" +
				"</body>" +
				"</html>";

			string date = GetDate();

			string simpleHtml = "<html>" +
								"<head>" +
								"<title>A page created with an image of an html page on it</title>" +
								"<meta name=\"created\" content=\"" + date + "\" />" +
								"</head>" +
								"<body>" +
								"<h1>This is a page with an image of an html page on it.</h1>" +
								"<img data-render-src=\"name:" + embeddedPartName +
								"\" alt=\"A website screen grab\" />" +
								"</body>" +
								"</html>";

			var createMessage = new HttpRequestMessage(HttpMethod.Post, apiRoute + "sections/" + sectionId + "/pages")
			{
				Content = new MultipartFormDataContent
				{
					{new StringContent(simpleHtml, Encoding.UTF8, "text/html"), "Presentation"},
					{new StringContent(embeddedWebPage, Encoding.UTF8, "text/html"), embeddedPartName}
				}
			};

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await HttpUtils.TranslateResponse(response);
		}

		/// <summary>
		/// Create a page with an image of a URL on it.
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="sectionId"></param>
		/// <param name="provider"></param>
		/// <param name="apiRoute"></param>
		/// <remarks>Create page using a multipart/form-data content type</remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<ApiBaseResponse> CreatePageWithUrl(bool debug, string sectionId, AuthProvider provider, string apiRoute)
		{
			if (debug)
			{
				Debugger.Launch();
				Debugger.Break();
			}

			var client = new HttpClient();

			// Note: API only supports JSON return type.
			client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

			// Not adding the Authentication header would produce an unauthorized call and the API will return a 401
			client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer",
				await Auth.GetAuthToken(provider));

			string date = GetDate();
			string simpleHtml = @"<html>" +
								"<head>" +
								"<title>A page created with an image from a URL on it</title>" +
								"<meta name=\"created\" content=\"" + date + "\" />" +
								"</head>" +
								"<body>" +
								"<p>This is a page with an image of an html page rendered from a URL on it.</p>" +
								"<img data-render-src=\"http://www.onenote.com\" alt=\"An important web page\"/>" +
								"</body>" +
								"</html>";

			var createMessage = new HttpRequestMessage(HttpMethod.Post, apiRoute + "sections/" + sectionId + "/pages")
			{
				Content = new StringContent(simpleHtml, Encoding.UTF8, "text/html")
			};

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await HttpUtils.TranslateResponse(response);
		}

		/// <summary>
		/// Create a page with note tags in it.
		/// http://blogs.msdn.com/b/onenotedev/archive/2014/10/17/announcing-tag-support.aspx
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="sectionId"></param>
		/// <param name="provider"></param>
		/// <param name="apiRoute"></param>
		/// <remarks>Create page using a multipart/form-data content type</remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<ApiBaseResponse> CreatePageWithNoteTags(bool debug, string sectionId, AuthProvider provider, string apiRoute)
		{
			if (debug)
			{
				Debugger.Launch();
				Debugger.Break();
			}

			var client = new HttpClient();

			// Note: API only supports JSON return type.
			client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

			// Not adding the Authentication header would produce an unauthorized call and the API will return a 401
			client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer",
				await Auth.GetAuthToken(provider));

			string date = GetDate();
			string simpleHtml = @"<html>" +
								"<head>" +
								"<title data-tag=\"to-do:completed\">A page created with note tags</title>" +
								"<meta name=\"created\" content=\"" + date + "\" />" +
								"</head>" +
								"<body>" +
								"<h1 data-tag=\"important\">Paragraphs with predefined note tags</h1>" +
								"<p data-tag=\"to-do\">Paragraph with note tag to-do (data-tag=\"to-do\")</p>" +
								"<p data-tag=\"important\">Paragraph with note tag important (data-tag=\"important\")</p>" +
								"<p data-tag=\"question\">Paragraph with note tag question (data-tag=\"question\")</p>" +
								"<p data-tag=\"definition\">Paragraph with note tag definition (data-tag=\"definition\")</p>" +
								"<p data-tag=\"highlight\">Paragraph with note tag highlight (data-tag=\"contact\")</p>" +
								"<p data-tag=\"contact\">Paragraph with note tag contact (data-tag=\"contact\")</p>" +
								"<p data-tag=\"address\">Paragraph with note tag address (data-tag=\"address\")</p>" +
								"<p data-tag=\"phone-number\">Paragraph with note tag phone-number (data-tag=\"phone-number\")</p>" +
								"<p data-tag=\"web-site-to-visit\">Paragraph with note tag web-site-to-visit (data-tag=\"web-site-to-visit\")</p>" +
								"<p data-tag=\"idea\">Paragraph with note tag idea (data-tag=\"idea\")</p>" +
								"<p data-tag=\"password\">Paragraph with note tag password (data-tag=\"critical\")</p>" +
								"<p data-tag=\"critical\">Paragraph with note tag critical (data-tag=\"project-a\")</p>" +
								"<p data-tag=\"project-a\">Paragraph with note tag project-a (data-tag=\"project-b\")</p>" +
								"<p data-tag=\"project-b\">Paragraph with note tag project-b (data-tag=\"remember-for-later\")</p>" +
								"<p data-tag=\"remember-for-later\">Paragraph with note tag remember-for-later (data-tag=\"remember-for-later\")</p>" +
								"<p data-tag=\"movie-to-see\">Paragraph with note tag movie-to-see (data-tag=\"movie-to-see\")</p>" +
								"<p data-tag=\"book-to-read\">Paragraph with note tag book-to-read (data-tag=\"book-to-read\")</p>" +
								"<p data-tag=\"music-to-listen-to\">Paragraph with note tag music-to-listen-to (data-tag=\"music-to-listen-to\")</p>" +
								"<p data-tag=\"source-for-article\">Paragraph with note tag source-for-article (data-tag=\"source-for-article\")</p>" +
								"<p data-tag=\"remember-for-blog\">Paragraph with note tag remember-for-blog (data-tag=\"remember-for-blog\")</p>" +
								"<p data-tag=\"discuss-with-person-a\">Paragraph with note tag discuss-with-person-a (data-tag=\"discuss-with-person-a\")</p>" +
								"<p data-tag=\"discuss-with-person-b\">Paragraph with note tag discuss-with-person-b (data-tag=\"discuss-with-person-a\")</p>" +
								"<p data-tag=\"discuss-with-manager\">Paragraph with note tag discuss-with-manager (data-tag=\"discuss-with-manager\")</p>" +
								"<p data-tag=\"send-in-email\">Paragraph with note tag send-in-email (data-tag=\"send-in-email\")</p>" +
								"<p data-tag=\"schedule-meeting\">Paragraph with note tag schedule-meeting (data-tag=\"schedule-meeting\")</p>" +
								"<p data-tag=\"call-back\">Paragraph with note tag call-back (data-tag=\"call-back\")</p>" +
								"<p data-tag=\"to-do-priority-1\">Paragraph with note tag to-do-priority-1 (data-tag=\"to-do-priority-1\")</p>" +
								"<p data-tag=\"to-do-priority-2\">Paragraph with note tag to-do-priority-2 (data-tag=\"to-do-priority-2\")</p>" +
								"<p data-tag=\"client-request\">Paragraph with note tag client-request (data-tag=\"client-request\")</p>" +
								"<br/>" +
								"<p style=\"font-size: 16px; font-family: Calibri, sans-serif\">Paragraphs with note tag status</p>" +
								"<p data-tag=\"to-do:completed\">Paragraph with note tag status completed</p>" +
								"<p data-tag=\"call-back:completed\">Paragraph with note tag status completed</p>" +
								"<br/>" +
								"<p style=\"font-size: 16px; font-family: Calibri, sans-serif\">Paragraph with multiple note tags</p>" +
								"<p data-tag=\"critical, question\">Paragraph with two note tags</p>" +
								"<p data-tag=\"password, send-in-email\">Multiple note tags</p>" +
								"<h1>List Item with a note tag</h1>" +
								"<li data-tag=\"to-do\" id=\"todoitem2\">Build a todo app with OneNote APIs</li>" +
								"<p style=\"font-size: 16px; font-family: Calibri, sans-serif\">Image with note tag</p>" +
								"<img data-tag=\"important\" src=\"http://placecorgi.com/300\" />" +
								"</body>" +
								"</html>";

			var createMessage = new HttpRequestMessage(HttpMethod.Post, apiRoute + "sections/" + sectionId + "/pages")
			{
				Content = new StringContent(simpleHtml, Encoding.UTF8, "text/html")
			};

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await HttpUtils.TranslateResponse(response);
		}

		#region Examples of POST https://www.onenote.com/api/v1.0/pages with auto-extraction of entities

		/// <summary>
		/// Create a page with auto-extracted business card on it.
		/// http://blogs.msdn.com/b/onenotedev/archive/2014/12/08/new-api-to-auto-extract-business-cards-recipe-urls-and-product-urls-now-in-beta.aspx
		/// Currently we can extract business cards from images and recipes, product info from urls.
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="sectionId"></param>
		/// <param name="provider"></param>
		/// <param name="apiRoute"></param>
		/// <remarks>Create page using a multipart/form-data content type</remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<ApiBaseResponse> CreatePageWithAutoExtractBusinessCard(bool debug, string sectionId, AuthProvider provider, string apiRoute)
		{
			if (debug)
			{
				Debugger.Launch();
				Debugger.Break();
			}

			var client = new HttpClient();

			// Note: API only supports JSON return type.
			client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

			// Not adding the Authentication header would produce an unauthorized call and the API will return a 401
			client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer",
				await Auth.GetAuthToken(provider));

			const string imagePartName = "image1";
			string date = GetDate();
			string simpleHtml = "<html>" +
								"<head>" +
								"<title>A page with auto-extracted business card from an image</title>" +
								"<meta name=\"created\" content=\"" + date + "\" />" +
								"</head>" +
								"<body>" +
								"<h1>This is a page with an extracted business card from an image</h1>" +
								"<div data-render-method=\"extract\" + " +
									"data-render-src=\"name:" + imagePartName +
									"\" data-render-fallback=\"none\" />" +
								"<p> This is the original image from which the business card was extracted: </p>" +
								"<img src=\"name:" + imagePartName + "\" />" +
								"</body>" +
								"</html>";

			// Create the image part - make sure it is disposed after we've sent the message in order to close the stream.
			HttpResponseMessage response;
			using (var imageContent = new StreamContent(await GetBinaryStream("assets\\BizCard.png")))
			{
				imageContent.Headers.ContentType = new MediaTypeHeaderValue("image/jpeg");
				var createMessage = new HttpRequestMessage(HttpMethod.Post, apiRoute + "sections/" + sectionId + "/pages")
				{
					Content = new MultipartFormDataContent
					{
						{new StringContent(simpleHtml, Encoding.UTF8, "text/html"), "Presentation"},
						{imageContent, imagePartName}
					}
				};

				// Must send the request within the using block, or the image stream will have been disposed.
				response = await client.SendAsync(createMessage);
			}

			return await HttpUtils.TranslateResponse(response);
		}

		/// <summary>
		/// Create a page with auto-extracted recipe on it.
		/// http://blogs.msdn.com/b/onenotedev/archive/2014/12/08/new-api-to-auto-extract-business-cards-recipe-urls-and-product-urls-now-in-beta.aspx
		/// Currently we can extract business cards from images and recipes, product info from urls.
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="sectionId"></param>
		/// <param name="provider"></param>
		/// <param name="apiRoute"></param>
		/// <remarks>Create page using a multipart/form-data content type</remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<ApiBaseResponse> CreatePageWithAutoExtractRecipe(bool debug, string sectionId, AuthProvider provider, string apiRoute)
		{
			if (debug)
			{
				Debugger.Launch();
				Debugger.Break();
			}

			var client = new HttpClient();

			// Note: API only supports JSON return type.
			client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

			// Not adding the Authentication header would produce an unauthorized call and the API will return a 401
			client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer",
				await Auth.GetAuthToken(provider));

			string date = GetDate();
			string simpleHtml = "<html>" +
			                    "<head>" +
			                    "<title>A page with auto-extracted recipe from a URL</title>" +
			                    "<meta name=\"created\" content=\"" + date + "\" />" +
			                    "</head>" +
			                    "<body>" +
			                    "<h1>This is a page with an extracted recipe from a URL: http://allrecipes.com/Recipe/Homemade-Mac-and-Cheese </h1>" +
			                    "<div data-render-method=\"extract\" + " +
			                    "data-render-src=\"http://allrecipes.com/Recipe/Homemade-Mac-and-Cheese\"" +
			                    "data-render-fallback=\"none\" />" +
			                    "<p> This is a screenshot of the original URL from which the product was extracted: </p>" +
			                    "<img data-render-src=\"http://allrecipes.com/Recipe/Homemade-Mac-and-Cheese\" />" +
			                    "</body>" +
			                    "</html>";

			var createMessage = new HttpRequestMessage(HttpMethod.Post, apiRoute + "sections/" + sectionId + "/pages")
			{
				Content = new StringContent(simpleHtml, Encoding.UTF8, "text/html")
			};

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await HttpUtils.TranslateResponse(response);
		}

		/// <summary>
		/// Create a page with auto-extracted product info on it.
		/// http://blogs.msdn.com/b/onenotedev/archive/2014/12/08/new-api-to-auto-extract-business-cards-recipe-urls-and-product-urls-now-in-beta.aspx
		/// Currently we can extract business cards from images and recipes, product info from urls.
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="sectionId"></param>
		/// <param name="provider"></param>
		/// <param name="apiRoute"></param>
		/// <remarks>Create page using a multipart/form-data content type</remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<ApiBaseResponse> CreatePageWithAutoExtractProduct(bool debug, string sectionId, AuthProvider provider, string apiRoute)
		{
			if (debug)
			{
				Debugger.Launch();
				Debugger.Break();
			}

			var client = new HttpClient();

			// Note: API only supports JSON return type.
			client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

			// Not adding the Authentication header would produce an unauthorized call and the API will return a 401
			client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer",
				await Auth.GetAuthToken(provider));

			string date = GetDate();
			string simpleHtml = "<html>" +
			                    "<head>" +
			                    "<title>A page with auto-extracted product info from a URL</title>" +
			                    "<meta name=\"created\" content=\"" + date + "\" />" +
			                    "</head>" +
			                    "<body>" +
			                    "<h1>This is a page with an extracted product info from a URL: http://www.amazon.com/Xbox-One/dp/B00KAI3KW2 </h1>" +
			                    "<div data-render-method=\"extract\" + " +
			                    "data-render-src=\"http://www.amazon.com/Xbox-One/dp/B00KAI3KW2\"" +
			                    "data-render-fallback=\"none\" />" +
			                    "<p> This is a screenshot of the original URL from which the recipe was extracted: </p>" +
			                    "<img data-render-src=\"http://www.amazon.com/Xbox-One/dp/B00KAI3KW2\" />" +
			                    "</body>" +
			                    "</html>";

			var createMessage = new HttpRequestMessage(HttpMethod.Post, apiRoute + "sections/" + sectionId + "/pages")
			{
				Content = new StringContent(simpleHtml, Encoding.UTF8, "text/html")
			};

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await HttpUtils.TranslateResponse(response);
		}

		#endregion

		#endregion

		#region Example of POST https://www.onenote.com/api/v1.0/pages?sectionName={AppName}

		/// <summary>
		/// Create a very simple page with some formatted text in a  specific section (name provided by caller)
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="sectionName">name of the section (in the default notebook) where the page will be created.
		///     If the section doesn't exist, this page create call will implicitly create the named section.</param>
		/// <param name="provider"></param>
		/// <param name="apiRoute"></param>
		/// <remarks>Create page using a single part text/html content type.
		/// This is a quick way to specify a custom sectionName that your app can save pages to.
		/// It uses the optional 'sectionName' query param in the pages endpoint.
		/// (e.g. saved all pages via my app under the '[AppName]' section).
		/// To fully control which notebook/section to create pages under; see the example below.
		/// NOTE: Using this approach, you can still create pages with ALL the different contents shown in examples above.
		/// </remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<ApiBaseResponse> CreateSimplePageInAGivenSectionName(bool debug, string sectionName, AuthProvider provider, string apiRoute)
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

			string date = GetDate();
			string simpleHtml = "<html>" +
								"<head>" +
								"<title>A simple page created under a named section</title>" +
								"<meta name=\"created\" content=\"" + date + "\" />" +
								"</head>" +
								"<body>" +
								"<p>This is a page that just contains some simple <i>formatted</i> <b>text</b></p>" +
								"<p>Here is a <a href=\"http://www.microsoft.com\">link</a></p>" +
								"</body>" +
								"</html>";

			// Prepare an HTTP POST request to the Pages endpoint
			// The request body content type is text/html
			var createMessage = new HttpRequestMessage(HttpMethod.Post, apiRoute + "pages?sectionName=" + WebUtility.UrlEncode(sectionName))
			{
				Content = new StringContent(simpleHtml, Encoding.UTF8, "text/html")
			};

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await HttpUtils.TranslateResponse(response);
		}

		#endregion

		#region Example of POST https://www.onenote.com/api/v1.0/sections/{sectionId}/pages

		/// <summary>
		/// Create a very simple page with some formatted text in a specific section
		/// </summary>
		/// <param name="debug">Run the code under the debugger</param>
		/// <param name="sectionId">Id of the section under which the page will be created</param>
		/// <param name="provider"></param>
		/// <param name="apiRoute"></param>
		/// <remarks>Create page using a single part text/html content type.
		/// The sectionId can be fetched by querying the user's sections (e.g. GET https://www.onenote.com/api/v1.0/sections ).
		/// NOTE: Using this approach, you can still create pages with ALL the different contents shown in examples above.
		/// </remarks>
		/// <returns>The converted HTTP response message</returns>
		public static async Task<ApiBaseResponse> CreateSimplePageInAGivenSectionId(bool debug, string sectionId, AuthProvider provider, string apiRoute)
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

			string date = GetDate();
			string simpleHtml = "<html>" +
								"<head>" +
								"<title>A simple page created under a specific notebook, section</title>" +
								"<meta name=\"created\" content=\"" + date + "\" />" +
								"</head>" +
								"<body>" +
								"<p>This is a page that just contains some simple <i>formatted</i> <b>text</b></p>" +
								"<p>Here is a <a href=\"http://www.microsoft.com\">link</a></p>" +
								"</body>" +
								"</html>";

			// Prepare an HTTP POST request to the Pages endpoint
			// The request body content type is text/html
			var createMessage = new HttpRequestMessage(HttpMethod.Post, apiRoute + "sections/" + sectionId + "/pages")
			{
				Content = new StringContent(simpleHtml, Encoding.UTF8, "text/html")
			};

			HttpResponseMessage response = await client.SendAsync(createMessage);

			return await HttpUtils.TranslateResponse(response);
		}
		#endregion

		#region Helper methods used in the examples

		/// <summary>
		/// Get date in ISO8601 format with local timezone offset
		/// </summary>
		/// <returns>Date as ISO8601 string</returns>
		private static string GetDate()
		{
			return DateTime.Now.ToString("o");
		}

		/// <summary>
		/// Get a binary file asset packaged with the application and return it as a managed stream
		/// </summary>
		/// <param name="binaryFile">The path to refer to the file relative to the application package root</param>
		/// <returns>A managed stream of the file data, opened for read</returns>
		private static async Task<Stream> GetBinaryStream(string binaryFile)
		{
			var storageFile = await Package.Current.InstalledLocation.GetFileAsync(binaryFile);
			var storageStream = await storageFile.OpenSequentialReadAsync();
			return storageStream.AsStreamForRead();
		}

		#endregion
	}
}
