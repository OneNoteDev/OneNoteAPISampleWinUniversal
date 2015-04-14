using OneNoteServiceSamplesWinUniversal.Data;
using OneNoteServiceSamplesWinUniversal.OneNoteApi;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Net;
using System.Text;
using Windows.UI.Xaml;

namespace OneNoteServiceSamplesWinUniversal.DataModel
{
	public class ItemPageModel : INotifyPropertyChanged
	{
		private string _authUserName = null;  // Logged in user name
		private SampleDataItem _item = null;  // Info about the current item
		private object _apiResponse = null;  // Response from an API request - Either a ApiBaseResponse or a List<ApiBaseResponse>
		private ApiBaseResponse _selectedResponse = null;  // Holds a reference to the currently selected response

		public event PropertyChangedEventHandler PropertyChanged;  // Implements INotifyPropertyChanged for notifying the binding engine when an attribute changes

		/// <summary>
		// NotifyPropertyChanged will fire the PropertyChanged event, 
		// passing the source property that is being updated.
		/// </summary>
		public void NotifyPropertyChanged(string propertyName)
		{
			if (PropertyChanged != null)
			{
				PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
			}
		}

		public string AuthUserName
		{
			get
			{
				return _authUserName;
			}

			set
			{
				_authUserName = value;
				NotifyPropertyChanged("AuthUserName");
				NotifyPropertyChanged("SignInButtonTitle");
			}
		}

		public string SignInButtonTitle
		{
			get
			{
				string title = "Sign In";

				if (Auth.IsSignedIn())
				{
					title = "Sign Out";
				}

				return title;
			}
		}

		public SampleDataItem Item
		{
			get
			{
				return _item;
			}

			set
			{
				_item = value;
				NotifyPropertyChanged("Item");
			}
		}

		public object ApiResponse
		{
			set
			{
				_apiResponse = value;

				// Update _selectedResponse as the first response in a list (or the only one returned, depends on the response type)
				ApiBaseResponse selectedResponse = null;

				if (_apiResponse != null)
				{
					if (_apiResponse is ApiBaseResponse)
					{
						selectedResponse = _apiResponse as ApiBaseResponse;
					}
					else
					{
						List<ApiBaseResponse> multipleResponses = _apiResponse as List<ApiBaseResponse>;
						selectedResponse = multipleResponses[0];
					}
				}

				SelectedResponse = selectedResponse;
				NotifyPropertyChanged("StatusCode");
				NotifyPropertyChanged("Body");
				NotifyPropertyChanged("BodyBorderThikness");
				NotifyPropertyChanged("IsResponseAvailable");
				NotifyPropertyChanged("IsResponseListAvailable");
				NotifyPropertyChanged("IsClientUrlAvailable");
				NotifyPropertyChanged("ResponseList");
				NotifyPropertyChanged("SelectedResponse");
			}
		}

		public ApiBaseResponse SelectedResponse
		{
			get
			{
				return _selectedResponse;
			}

			set
			{
				_selectedResponse = value;
				NotifyPropertyChanged("OneNoteClientUrl");
				NotifyPropertyChanged("OneNoteWebUrl");
			}
		}

		public string StatusCode
		{
			get
			{
				string statusCodeString = string.Empty;

				if (_selectedResponse != null)
				{
					HttpStatusCode statusCode = _selectedResponse.StatusCode;

					if (statusCode != 0)
					{
						statusCodeString = string.Format("{0}-{1}", (int)statusCode, statusCode.ToString());
					}
				}

				return statusCodeString;
			}
		}

		public string Body
		{
			get
			{
				string body = string.Empty;

				if (_selectedResponse != null && _selectedResponse.Body != null)
				{
					body = _selectedResponse.Body;
				}

				return body;
			}
		}

		public Thickness BodyBorderThikness
		{
			get
			{
				int bodyBorderThikness = 2;

				if (_selectedResponse != null)
				{
					bodyBorderThikness = 0;
				}

				return new Thickness(bodyBorderThikness);
			}
		}

		public bool IsResponseAvailable
		{
			get
			{
				return _apiResponse != null;
			}
		}

		public Visibility IsResponseListAvailable
		{
			get
			{
				Visibility responseListVisible = Visibility.Collapsed;

				if (_apiResponse is List<ApiBaseResponse>)
				{
					responseListVisible = Visibility.Visible;
				}

				return responseListVisible;
			}
		}

		public List<ApiBaseResponse> ResponseList
		{
			get
			{
				return _apiResponse as List<ApiBaseResponse>;
			}
		}

		public Visibility IsClientUrlAvailable
		{
			get
			{
				Visibility urlVisible = Visibility.Collapsed;

				if (!String.IsNullOrEmpty(OneNoteClientUrl))
				{
					urlVisible = Visibility.Visible;
				}

				return urlVisible;
			}
		}

		public string OneNoteClientUrl
		{
			get
			{
				string clientUrl = string.Empty;

				if (_selectedResponse != null && _selectedResponse.Links != null)
				{
					clientUrl = _selectedResponse.Links.OneNoteClientUrl.Href;
				}

				return clientUrl;
			}
		}

		public string OneNoteWebUrl
		{
			get
			{
				string webUrl = string.Empty;

				if (_selectedResponse != null && _selectedResponse.Links != null)
				{
					webUrl = _selectedResponse.Links.OneNoteWebUrl.Href;
				}

				return webUrl;
			}
		}
	}
}
