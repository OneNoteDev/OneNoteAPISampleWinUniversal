
## OneNote service API Windows Universal Sample README

Created by Microsoft Corporation, 2014. Provided As-is without warranty. Trademarks mentioned here are the property of their owners.

### Intro
* [Universal Windows apps](http://blogs.windows.com/buildingapps/2014/04/02/extending-platform-commonality-through-universal-windows-apps/)
* This is a **newer and better** version of the previously released oneNote API [WinStore](https://github.com/OneNoteDev/OneNoteAPISampleWinStore) and [WinPhone](https://github.com/OneNoteDev/OneNoteAPISampleWinPhone) code samples. Use this code sample to build universal Windows 8.1 and above apps.
* As of December 2014, This code sample contains examples of all shipped features as well as most of the beta features released by the OneNote API team.

### API functionality demonstrated in this sample

The following aspects of the API are covered in this sample. You can find additional documentation at the links below.

* [Authenticate the user using the OnlineIdAuthenticator API](http://msdn.microsoft.com/en-us/library/windows/apps/windows.security.authentication.onlineid.onlineidauthenticator.aspx)
    * In our previous [WinStore](https://github.com/OneNoteDev/OneNoteAPISampleWinStore) and [WinPhone](https://github.com/OneNoteDev/OneNoteAPISampleWinPhone) Code samples we demonstrated how to use the 
[LiveSDK](http://msdn.microsoft.com/EN-US/library/office/dn575435.aspx) to do OAuth against the Microsoft Account service. This code sample, demonstrates an alternative way to do OAuth using the new Windows.Security.Authentication.OnlineId.OnlineIdAuthenticator class. Both the existing Live SDK approach and this alternative will work in Windows 8.1 universal apps
    * The usage of the OnlineIdAuthenticator class is based on the [Windows universal code sample](http://code.msdn.microsoft.com/windowsapps/Windows-account-authorizati-7c95e284)
* Create Pages: 
    * [POST simple HTML to a new OneNote QuickNotes page](http://msdn.microsoft.com/EN-US/library/office/dn575428.aspx)
    * [POST multi-part message with image data included in the request](http://msdn.microsoft.com/EN-US/library/office/dn575432.aspx)
    * POST page with a snapshot of an embedded web page HTML
    * [POST page with a URL rendered as an image](http://msdn.microsoft.com/EN-US/library/office/dn575431.aspx)
    * [POST page with a file attachment](http://msdn.microsoft.com/en-us/library/office/dn575436.aspx)
    * [POST page with a PDF file rendered and attached](http://msdn.microsoft.com/EN-US/library/office/dn655137.aspx)
    * POST page with note tags
    * POST page with a business card automatically extracted from an image
    * POST page with a recipe automatically extracted from a URL
    * POST page with a product info automatically extracted from a URL
    * [Extract the returned oneNoteClientURL and oneNoteWebURL links](http://msdn.microsoft.com/EN-US/library/office/dn575433.aspx)
    * [POST page in a specific named section](http://msdn.microsoft.com/EN-US/library/office/dn672416.aspx)
    * POST page under a specific notebook and section
* Query and Search Pages:
    *  GET a paginated list of all pages in OneNote
    *  GET metadata for a specific page
    *  GET pages with title containing a specifc substring using the $filter query parameter and contains method
    *  GET pages using [OData v4](http://docs.oasis-open.org/odata/odata/v4.0/os/part1-protocol/odata-v4.0-os-part1-protocol.html) query parameters like $skip and $top
    *  GET a sorted list of pages using the $orderBy query parameter
    *  GET selected metadata for pages using the $select query parameter
    *  GET pages containing the matching search term using the search query parameter
    *  GET back a specific page's content as HTML
* Manage Notebooks and Sections:
    * GET all notebooks and sections in one roundtrip using the $expand query parameter
    * GET a list of all notebooks
    * GET metadata for a specific notebook
    * GET a list of all sections
    * GET metadata for a specific section
    * GET a list of all SectionGroups
    * GET notebooks and sections with a specific name using the $filter query parameter
    * GET notebooks shared by others using the $filter query parameter and useRole property
    * GET a sorted list of notebooks using the $orderBy query parameter
    * GET selected metadata for notebooks using the $select query parameter
    * GET a list of all sections under a specific notebook
    * POST a new notebook
    * POST a new section under a specific notebook
* Updates Pages:
    * PATCH an existing page and append new content to it
    * PATCH an existing page and prepend new content to it
    * PATCH an existing page and replace an existing image/attachment with a new one

If you wish to ignore the app UI/design and directly get to the code that interacts with the OneNote API, you can find them  under the .../OneNoteServiceSamplesWinUniversal.Shared/OneNoteApi folder. 

### Prerequisites

**Tools and Libraries** you will need to download, install, and configure for your development environment. 
* Operating system requirements for developing Windows universal apps: 
    *   Client: Windows 8.1 OR 
    *   Server: Windows Server 2012 R2 
    *   Phone: Windows Phone 8.1 (developer unlocked phone for device testing. Alternatively emulators can be used)
* [Visual Studio 2013 Update 2 or later](http://www.visualstudio.com/en-us/downloads). 

Be sure you enable the Windows Phone SDK when you install Visual Studio. 
If you don't, you will need to uninstall and re-install Visual Studio to get those features.

* You can get a full list of tools needed to build Universal Windows app [here](http://dev.windows.com/en-us/develop/downloads)

* **NuGet packages** used in the sample. These are handled using the package 
manager, as described in the setup instructions. These should update 
automatically at build time; if not, make sure your NuGet package manager 
is up-to-date. You can learn more about the packages we used at the links below.
    * [Newtonsoft Json.NET package](http://newtonsoft.com/) provides Json parsing utilities.

**Accounts**

* As the developer, you'll need to [have a Microsoft account and get a client ID string](http://msdn.microsoft.com/EN-US/library/office/dn575426.aspx) so your app can authenticate.

### Using the sample

After you've setup your development tools, and installed the prerequisites listed above,...

1. Download the repo as a ZIP file to your local computer, and extract the files. Or, clone the repository into a local copy of Git.
2. Open the project in Visual Studio.
3. Get a client ID string and copy it into the file .../OneNoteServiceSamplesWinUniversal.Shared/OneNoteApi/Auth.cs (~line 54).
4. Build and run the app (F5). 

   (If your copy of NuGet is up-to-date, it should automatically 
update the packages. If you get package-not-found errors, update NuGet and rebuild, and that 
should fix it.)

5. Sign in to your Microsoft account in the running app.
6. Allow the app to exercise the desired OneNote API call.
7. If you wish to use the code sample in your own app, be sure to [Create an app package and use appropriate certificates](http://docs.oasis-open.org/odata/odata/v4.0/os/part1-protocol/odata-v4.0-os-part1-protocol.html)
 
### Version Info

This is the initial public release for this code sample.
  
### Learning More

* Visit the [dev.onenote.com](http://dev.onenote.com) Dev Center
* Contact us on [StackOverflow (tagged OneNote)](http://go.microsoft.com/fwlink/?LinkID=390182)
* Follow us on [Twitter @onenotedev](http://www.twitter.com/onenotedev)
* Read our [OneNote Developer blog](http://go.microsoft.com/fwlink/?LinkID=390183)
* Explore the API using the [apigee.com interactive console](http://go.microsoft.com/fwlink/?LinkID=392871).
Also, see the [short overview/tutorial](http://go.microsoft.com/fwlink/?LinkID=390179). 
* [API Reference](http://msdn.microsoft.com/en-us/library/office/dn575437.aspx) documentation
* [Debugging / Troubleshooting](http://msdn.microsoft.com/EN-US/library/office/dn575430.aspx)
* [Getting Started](http://go.microsoft.com/fwlink/?LinkID=331026) with the OneNote service API
