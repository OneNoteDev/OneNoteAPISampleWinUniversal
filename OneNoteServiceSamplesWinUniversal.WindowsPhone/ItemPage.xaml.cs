using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Windows.ApplicationModel.Core;
using Windows.Foundation;
using Windows.System;
using Windows.UI.Core;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Navigation;
using OneNoteServiceSamplesWinUniversal.Common;
using OneNoteServiceSamplesWinUniversal.Data;
using OneNoteServiceSamplesWinUniversal.OneNoteApi;
using OneNoteServiceSamplesWinUniversal.DataModel;

// The Universal Hub Application project template is documented at http://go.microsoft.com/fwlink/?LinkID=391955

namespace OneNoteServiceSamplesWinUniversal
{
	/// <summary>
	/// A page that displays details for a single item within a group.
	/// </summary>
	public sealed partial class ItemPage : SharedBasePage
	{
		private readonly NavigationHelper _navigationHelper;
        private readonly ItemPageModel _model = new ItemPageModel();

		public ItemPage()
		{
			InitializeComponent();
            DataContext = Model;
			_navigationHelper = new NavigationHelper(this);
			_navigationHelper.LoadState += NavigationHelper_LoadState;
			_navigationHelper.SaveState += NavigationHelper_SaveState;
		} 

		/// <summary>
		/// Gets the <see cref="NavigationHelper"/> associated with this <see cref="Page"/>.
		/// </summary>
		public NavigationHelper NavigationHelper
		{
			get { return _navigationHelper; }
		}

        public ItemPageModel Model
        {
            get { return _model; }
        }

		/// <summary>
		/// Populates the page with content passed during navigation.  Any saved state is also
		/// provided when recreating a page from a prior session.
		/// </summary>
		/// <param name="sender">
		/// The source of the event; typically <see cref="NavigationHelper"/>
		/// </param>
		/// <param name="e">Event data that provides both the navigation parameter passed to
		/// <see cref="Frame.Navigate(Type, object)"/> when this page was initially requested and
		/// a dictionary of state preserved by this page during an earlier
		/// session.  The state will be null the first time a page is visited.</param>
		private async void NavigationHelper_LoadState(object sender, LoadStateEventArgs e)
		{
			var item = await SampleDataSource.GetItemAsync((string)e.NavigationParameter);
            Model.Item = item;
            InputSelectionPanel2.Visibility = (item.RequiresInputComboBox2) ? Visibility.Visible : Visibility.Collapsed;
			InputTextBox.Visibility = (item.RequiresInputTextBox) ? Visibility.Visible : Visibility.Collapsed;
			if (item.RequiresInputComboBox1)
			{
				var response = await SampleDataSource.ExecuteApiPrereq(item.UniqueId, Auth.GetAuthToken(UserData.Provider).Result);
				if (response is List<ApiBaseResponse>)
				{
					InputComboBox1.ItemsSource = response;
					InputSelectionPanel1.Visibility = Visibility.Visible;
				}
			}
			else
			{
				InputSelectionPanel1.Visibility = Visibility.Collapsed;
			}
		}

		/// <summary>
		/// Preserves state associated with this page in case the application is suspended or the
		/// page is discarded from the navigation cache.  Values must conform to the serialization
		/// requirements of <see cref="SuspensionManager.SessionState"/>.
		/// </summary>
		/// <param name="sender">The source of the event; typically <see cref="NavigationHelper"/></param>
		/// <param name="e">Event data that provides an empty dictionary to be populated with
		/// serializable state.</param>
		private void NavigationHelper_SaveState(object sender, SaveStateEventArgs e)
		{
			// TODO: Save the unique state of the page here.
		}

		#region NavigationHelper registration

		/// <summary>
		/// The methods provided in this section are simply used to allow
		/// NavigationHelper to respond to the page's navigation methods.
		/// <para>
		/// The navigation parameter is available in the LoadState method
		/// in addition to page state preserved during an earlier session.
		/// </para>
		/// </summary>
		/// <param name="e">Event data that describes how this page was reached.</param>
		protected async override void OnNavigatedTo(NavigationEventArgs e)
		{
			_navigationHelper.OnNavigatedTo(e);
            Model.AuthUserName = await Auth.GetUserName();
        }

		protected override void OnNavigatedFrom(NavigationEventArgs e)
		{
			_navigationHelper.OnNavigatedFrom(e);
		}

		private async void Button_Click(object sender, RoutedEventArgs e)
		{
			RunButton.IsEnabled = false;
            SampleDataItem item = Model.Item;

			await ExecuteApiAction(false, item);
			RunButton.IsEnabled = true;
		}

		private async Task ExecuteApiAction(bool debug, SampleDataItem item)
		{
            Model.ApiResponse = null;

			string requiredSelectedId = null;
			if (item.RequiresInputComboBox1)
			{
				var selectedItem = (ApiBaseResponse)InputComboBox1.SelectedItem;
				if (selectedItem != null)
				{
					requiredSelectedId = selectedItem.Id;
				}
				else
				{
					FlyoutBase.ShowAttachedFlyout(InputComboBox1);
					return;
				}
			}
			if (item.RequiresInputComboBox2)
			{
				var selectedItem = (ApiBaseResponse)InputComboBox2.SelectedItem;
				if (selectedItem != null)
				{
					requiredSelectedId = selectedItem.Id;
				}
				else
				{
					FlyoutBase.ShowAttachedFlyout(InputComboBox2);
					return;
				}
			}

			string requiredInputText = null;
			if (item.RequiresInputTextBox)
			{
				requiredInputText = InputTextBox.Text;
				if (String.IsNullOrWhiteSpace(requiredInputText) || requiredInputText.Equals("Enter required text here"))
				{
					FlyoutBase.ShowAttachedFlyout(InputTextBox);
					return;
				}
			}
			UserData.TimeStamp = DateTime.UtcNow;
			Model.UserData = UserData;
            Model.ApiResponse = await SampleDataSource.ExecuteApi(item.UniqueId, debug, requiredSelectedId, requiredInputText, Auth.GetAuthToken(_navigationHelper.Provider).Result);
            Model.AuthUserName = await Auth.GetUserName();
		}

		private void SetResponseLinks(ApiBaseResponse apiBaseResponse)
		{
			if (apiBaseResponse != null && apiBaseResponse.Links != null)
			{
				ClientLinkLaunchButton.DataContext = apiBaseResponse.Links.OneNoteClientUrl.Href;
			}
		}

		/// <summary>
		/// Yield the thread of operation to the dispatcher to allow UI updates to happen.
		/// </summary>
		/// <remarks>
		/// Schedules a do-nothing operation on the dispatcher, then allows continuation after it's 'completed'.
		/// Usage: await Yield();
		/// </remarks>
		/// <returns>A dispatcher operation that is awaitable.</returns>
		/// 
		private static IAsyncAction Yield()
		{
			return CoreApplication.MainView.CoreWindow.Dispatcher.RunAsync(CoreDispatcherPriority.Normal, () => { });
		}

		/// <summary>
		/// Update the enabled states of the output response UI controls and allow the UI to refresh.
		/// </summary>
		private async Task SetOutputResponseControlsState(bool enabled)
		{
			var childrenOfOutputGrid = OutputRegionGrid.Children.OfType<Control>();
			foreach (var child in childrenOfOutputGrid)
			{
				child.IsEnabled = enabled;
			}
			await Yield();
		}

		private void InputTextBox_OnGotFocus(object sender, RoutedEventArgs e)
		{
			//Clear the section hint if that was the existing value
			string sectionSpecified = InputTextBox.Text;
			if (sectionSpecified.Equals("Enter required text here"))
			{
				InputTextBox.Text = string.Empty;
			}
		}

		private void InputTextBox_OnLostFocus(object sender, RoutedEventArgs e)
		{
			//Restore the section hint if a value was not entered
			string sectionSpecified = InputTextBox.Text;
			if (string.IsNullOrEmpty(sectionSpecified))
			{
				InputTextBox.Text = "Enter required text here";
			}
		}

		private async void ClientLinkLaunchButton_OnClickLinkLaunchButton_Click(object sender, RoutedEventArgs e)
		{
			Uri uri;
			string link = ClientLinkLaunchButton.Tag as string;
			if (Uri.TryCreate(link, UriKind.Absolute, out uri))
			{
				await Launcher.LaunchUriAsync(uri);
			}
		}

		private void ResponseListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			var comboBox = sender as ComboBox;
			if (comboBox != null && comboBox.SelectedItem != null)
			{
				SetResponseLinks(comboBox.SelectedItem as ApiBaseResponse);
			}
		}

		private void InputComboBox1_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (InputSelectionPanel2.Visibility == Visibility.Visible)
			{
				var comboBox = sender as ComboBox;
				if (comboBox != null && comboBox.SelectedItem != null)
				{
					var response = comboBox.SelectedItem as GenericEntityResponse;
					if (response != null && response.Sections != null)
					{
						InputComboBox2.ItemsSource = response.Sections;
					}
				}
			}
		}
		#endregion
	}
}