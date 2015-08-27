
## OneNote service API Windows Universal Sample README

Created by Microsoft Corporation, 2015. Provided As-is without warranty. Trademarks mentioned here are the property of their owners.

### Intro
* [Universal Windows apps](http://blogs.windows.com/buildingapps/2014/04/02/extending-platform-commonality-through-universal-windows-apps/)
* This is a **newer and better** version of the previously released oneNote API [WinStore](https://github.com/OneNoteDev/OneNoteAPISampleWinStore) and [WinPhone](https://github.com/OneNoteDev/OneNoteAPISampleWinPhone) code samples. Use this code sample to build universal Windows 8.1 and above apps.
* As of August 2015, this code sample contains examples of all shipped features as well as most of the beta features released by the OneNote API team.

**If you wish to ignore the app UI/design and directly get to the part that interacts with the OneNote API, you can find the code under the [OneNoteServiceSamplesWinUniversal.Shared/OneNoteApi](https://github.com/OneNoteDev/OneNoteAPISampleWinUniversal/tree/master/OneNoteServiceSamplesWinUniversal.Shared/OneNoteApi) folder**. 

### OneDrive based API

You can find additional documentation at the links below.

* [Authenticate the user using the OnlineIdAuthenticator API](http://msdn.microsoft.com/en-us/library/windows/apps/windows.security.authentication.onlineid.onlineidauthenticator.aspx)
    * In our previous [WinStore](https://github.com/OneNoteDev/OneNoteAPISampleWinStore) and [WinPhone](https://github.com/OneNoteDev/OneNoteAPISampleWinPhone) Code samples we demonstrated how to use the 
[LiveSDK](http://msdn.microsoft.com/EN-US/library/office/dn575435.aspx) to do OAuth against the Microsoft Account service. This code sample, demonstrates an alternative way to do OAuth using the new Windows.Security.Authentication.OnlineId.OnlineIdAuthenticator class. Both the existing Live SDK approach and this alternative will work in Windows 8.1 universal apps
    * The usage of the OnlineIdAuthenticator class is based on the [Windows universal code sample](http://code.msdn.microsoft.com/windowsapps/Windows-account-authorizati-7c95e284)
* Create Pages: 
    * [POST simple HTML to a new OneNote Quick Notes page](http://msdn.microsoft.com/EN-US/library/office/dn575428.aspx)
    * [POST multi-part message with image data included in the request](http://msdn.microsoft.com/EN-US/library/office/dn575432.aspx)
    * [POST page with a snapshot of an embedded web page HTML](http://msdn.microsoft.com/EN-US/library/office/dn575431.aspx)
    * [POST page with a URL rendered as an image](http://msdn.microsoft.com/EN-US/library/office/dn575431.aspx)
    * [POST page with a file attachment](http://msdn.microsoft.com/en-us/library/office/dn575436.aspx)
    * [POST page with a PDF file rendered and attached](http://msdn.microsoft.com/EN-US/library/office/dn655137.aspx)
    * [POST page with note tags](https://msdn.microsoft.com/EN-US/library/office/mt159148.aspx)
    * [POST page with a business card automatically extracted from an image](https://msdn.microsoft.com/EN-US/library/office/dn600338.aspx#ExtractBusinessCard)
    * [POST page with a recipe automatically extracted from a URL](https://msdn.microsoft.com/EN-US/library/office/dn600338.aspx#ExtractRecipe)
    * [POST page with a product info automatically extracted from a URL](https://msdn.microsoft.com/EN-US/library/office/dn600338.aspx#ExtractProduct)
    * [Extract the returned oneNoteClientURL and oneNoteWebURL links](http://msdn.microsoft.com/EN-US/library/office/dn575433.aspx)
    * [POST page in a specific named section](http://msdn.microsoft.com/EN-US/library/office/dn672416.aspx)
    * [POST page under a specific notebook and section](http://dev.onenote.com/docs#/reference/post-pages/v10menotessectionsidpages/post)
* Query and Search Pages:
    *  [GET a paginated list of all pages in OneNote](http://dev.onenote.com/docs#/reference/get-pages)
    *  [GET metadata for a specific page](http://dev.onenote.com/docs#/reference/get-pages/v10menotespagesid/get)
    *  [GET pages with title containing a specific substring using the $filter query parameter and contains method](http://dev.onenote.com/docs#/reference/get-pages/v10menotespagesfilterorderbyselecttopskipsearchcount/get)
    *  [GET pages using OData v4 query parameters like $skip and $top](http://dev.onenote.com/docs#/reference/get-pages/v10menotespagesfilterorderbyselecttopskipsearchcount/get)
        * [OData v4](http://docs.oasis-open.org/odata/odata/v4.0/os/part1-protocol/odata-v4.0-os-part1-protocol.html)
    *  [GET a sorted list of pages using the $orderby query parameter](http://dev.onenote.com/docs#/reference/get-pages/v10menotespagesfilterorderbyselecttopskipsearchcount/get)
    *  [GET selected metadata for pages using the $select query parameter](http://dev.onenote.com/docs#/reference/get-pages/v10menotespagesfilterorderbyselecttopskipsearchcount/get)
    *  [GET pages containing the matching search term using the search query parameter](http://dev.onenote.com/docs#/reference/get-pages/v10menotespagesfilterorderbyselecttopskipsearchcount/get)
    *  [GET back a specific page's content as HTML](http://dev.onenote.com/docs#/reference/get-pages/v10menotespagesidcontentincludeids/get)
* Manage Notebooks and Sections:
    * GET all notebooks and sections in one round trip using the $expand query parameter
    * [GET a list of all notebooks](http://dev.onenote.com/docs#/reference/get-notebooks)
    * [GET metadata for a specific notebook](http://dev.onenote.com/docs#/reference/get-notebooks/v10menotesnotebooksidselectexpand/get)
    * [GET a list of all sections](http://dev.onenote.com/docs#/reference/get-sections)
    * [GET metadata for a specific section](http://dev.onenote.com/docs#/reference/get-sections/v10menotessectionsidselectexpand/get)
    * [GET a list of all section groups](http://dev.onenote.com/docs#/reference/get-sectiongroups)
    * [GET notebooks and sections with a specific name using the $filter query parameter](http://dev.onenote.com/docs#/reference/get-notebooks/v10menotesnotebooksfilterorderbyselectexpandtopskipcount/get)
    * [GET notebooks shared by others using the $filter query parameter and userRole property](http://dev.onenote.com/docs#/reference/get-notebooks/v10menotesnotebooksfilterorderbyselectexpandtopskipcount/get)
    * [GET a sorted list of notebooks using the $orderby query parameter](http://dev.onenote.com/docs#/reference/get-notebooks/v10menotesnotebooksfilterorderbyselectexpandtopskipcount/get)
    * [GET selected metadata for notebooks using the $select query parameter](http://dev.onenote.com/docs#/reference/get-notebooks/v10menotesnotebooksfilterorderbyselectexpandtopskipcount/get)
    * [GET a list of all sections under a specific notebook](http://dev.onenote.com/docs#/reference/get-notebooks/v10menotesnotebooksfilterorderbyselectexpandtopskipcount/get)
    * [POST a new notebook](http://dev.onenote.com/docs#/reference/post-notebooks)
    * [POST a new section group under a specific notebook](http://dev.onenote.com/docs#/reference/post-sectiongroups/betamenotesnotebooksidsectiongroups/post)
    * [POST a new section group under a specific section group](http://dev.onenote.com/docs#/reference/post-sectiongroups/betamenotessectiongroupsidsectiongroups/post)
    * [POST a new section under a specific notebook](http://dev.onenote.com/docs#/reference/post-sections/v10menotesnotebooksidsections/post)
    * [POST a new section under a specific section group](http://dev.onenote.com/docs#/reference/post-sections/betamenotessectiongroupsidsections/post)
* Update Pages:  
    * Examples of [PATCH Pages](http://dev.onenote.com/docs#/reference/patch-pages)  
* Delete Pages:  
    * [DELETE page](http://dev.onenote.com/docs#/reference/delete-pages)

### O365 based API (only available on beta)
All the APIs for Microsoft Account (LiveId) are supported in O365 as well except the following:  
* $search query option is currently not supported.


### Prerequisites

**Tools and Libraries** you will need to download, install, and configure for your development environment. 
* Operating system requirements for developing Windows universal apps: 
    *   Client: Windows 8.1 OR 
    *   Server: Windows Server 2012 R2 
    *   Phone: Windows Phone 8.1 (developer unlocked phone for device testing. Alternatively emulators can be used)
* [Visual Studio 2013 Update 4 or later](http://www.visualstudio.com/en-us/downloads). 

Be sure you enable the Windows Phone SDK when you install Visual Studio. 
If you don't, you will need to un-install and re-install Visual Studio to get those features.

For using emulator (Windows Phone), make sure to enable hyper-v

* You can get a full list of tools needed to build Universal Windows app [here](http://dev.windows.com/en-us/develop/downloads)

** O365 specific Pre-requisites
* What you need to do is to go to Programs and Features, Turn on/off Windows features and then choose to add Windows Identity Foundation 3.5 to your system.

* **NuGet packages** used in the sample. These are handled using the package 
manager, as described in the setup instructions. These should update 
automatically at build time; if not, make sure your NuGet package manager 
is up-to-date. You can learn more about the packages we used at the links below.
    * [Newtonsoft Json.NET package](http://newtonsoft.com/) provides Json parsing utilities.
    * [Active Directory Authentication Library package](https://www.nuget.org/packages/Microsoft.IdentityModel.Clients.ActiveDirectory/) provides AAD authentication utilities (for Office 365)

### Accounts
**To use this Code Sample in your own Microsoft Account based app, be sure to do the following:**
* As the developer, you'll need to [have a Microsoft account and get a client ID string](http://msdn.microsoft.com/EN-US/library/office/dn575426.aspx) so your app can authenticate. Get a client ID string and copy it into the file under [.../OneNoteServiceSamplesWinUniversal.Shared/OneNoteApi/LiveIdAuth.cs](https://github.com/OneNoteDev/OneNoteAPISampleWinUniversal/blob/master/OneNoteServiceSamplesWinUniversal.Shared/OneNoteApi/LiveIdAuth.cs#L54) (~line 54).
* Use the least permissible scopes by editing [.../OneNoteServiceSamplesWinUniversal.Shared/OneNoteApi/LiveIdAuth.cs](https://github.com/OneNoteDev/OneNoteAPISampleWinUniversal/blob/master/OneNoteServiceSamplesWinUniversal.Shared/OneNoteApi/LiveIdAuth.cs#L61) (~line 61).
* [Create an app package and use appropriate certificates](http://msdn.microsoft.com/en-us/library/windows/apps/xaml/hh975357.aspx)

**To use this Code Sample in your own Microsoft Office 365 based app, refer to the details in the blog [Support for work and school notebooks on Office 365](http://blogs.msdn.com/b/onenotedev/archive/2015/04/30/support-for-work-and-school-notebooks-on-office-365-in-preview.aspx)

### Using the sample

After you've set up your development tools, and installed the prerequisites listed above,...

1. Download the repository as a ZIP file to your local computer, and extract the files. Or, clone the repository into a local copy of Git.
2. Open the project (.sln file) in Visual Studio.
3. It is highly recommended that you get your own client ID string and copy it into the file under 
	[.../OneNoteServiceSamplesWinUniversal.Shared/OneNoteApi/LiveIdAuth.cs](https://github.com/OneNoteDev/OneNoteAPISampleWinUniversal/blob/master/OneNoteServiceSamplesWinUniversal.Shared/OneNoteApi/LiveIdAuth.cs#L54) (~line 54)
	OR
	[.../OneNoteServiceSamplesWinUniversal.Shared/OneNoteApi/o365Auth.cs](https://github.com/OneNoteDev/OneNoteAPISampleWinUniversal/blob/master/OneNoteServiceSamplesWinUniversal.Shared/OneNoteApi/o365Auth.cs#L40) (~line 40)
4. For o365, additionally, change RedirectUri to one of your app under
	[.../OneNoteServiceSamplesWinUniversal.Shared/OneNoteApi/o365Auth.cs](https://github.com/OneNoteDev/OneNoteAPISampleWinUniversal/blob/master/OneNoteServiceSamplesWinUniversal.Shared/OneNoteApi/o365Auth.cs#L52) (~line 52)

5. Build and run the application (F5). 

(If your copy of NuGet is up-to-date, it should automatically update the packages. If you get package-not-found errors, update NuGet and rebuild, and that should fix it.)

If you get any build or deployment errors related to the Publisher of the app (Microsoft Account), be sure to 
[Create LiveId app package and use appropriate certificates](http://msdn.microsoft.com/en-us/library/windows/apps/xaml/hh975357.aspx)

If you get any build or deployment errors related to the Publisher of the app (Microsoft Office 365), be sure to 
[Have your app appear in the Office 365 app launcher](https://msdn.microsoft.com/en-us/office/office365/howto/connect-your-app-to-o365-app-launcher)

### Version Info

| Date | Change |
|------|------|
| August 2015 | Added beta `POST /sectiongroups/{id}/sections`, `POST /notebooks/{id}/sectiongroups`, and  `POST /sectiongroups/{id}/sectiongroups` examples.<br/>Changed `DELETE /pages/{id}` from beta to v1.0.<br/>Support for `POST /pages` and `POST /pages?sectionName` in Office 365. |
| April 2015 | Added Office 365 with Azure AD auth support<br/>Added beta `DELETE /pages/{id}` example. |
| March 2015 | Added `PATCH` example |
| January 2015 | Added `GET /sections/{id}/pages` examples. |
| December 2014 | Initial public release for this code sample. |
  
### Learning More

* Visit the [dev.onenote.com](http://dev.onenote.com) Dev Center
* Contact us on [StackOverflow (tagged OneNote)](http://go.microsoft.com/fwlink/?LinkID=390182)
* Follow us on [Twitter @onenotedev](http://www.twitter.com/onenotedev)
* Read our [OneNote Developer blog](http://go.microsoft.com/fwlink/?LinkID=390183)
* Explore the API using the [apigee.com interactive console](http://go.microsoft.com/fwlink/?LinkID=392871) OR the [apiary.io interactive console](http://dev.onenote.com/docs).
Also, see the [short overview/tutorial](http://go.microsoft.com/fwlink/?LinkID=390179). 
* [API Reference](http://msdn.microsoft.com/en-us/library/office/dn575437.aspx) documentation
* [Debugging / Troubleshooting](http://msdn.microsoft.com/EN-US/library/office/dn575430.aspx)
* [Getting Started](http://go.microsoft.com/fwlink/?LinkID=331026) with the OneNote service API
