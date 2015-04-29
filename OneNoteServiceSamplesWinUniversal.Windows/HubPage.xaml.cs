using System;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Media.Imaging;
using Windows.UI.Xaml.Navigation;
using OneNoteServiceSamplesWinUniversal.Common;
using OneNoteServiceSamplesWinUniversal.Data;
using OneNoteServiceSamplesWinUniversal.OneNoteApi;

// The Universal Hub Application project template is documented at http://go.microsoft.com/fwlink/?LinkID=391955
namespace OneNoteServiceSamplesWinUniversal
{
	/// <summary>
	/// A page that displays a grouped collection of items.
	/// </summary>
	public sealed partial class HubPage : SharedBasePage
	{
		private readonly NavigationHelper _navigationHelper;
		private readonly ObservableDictionary _defaultViewModel = new ObservableDictionary();

		/// <summary>
		/// Gets the NavigationHelper used to aid in navigation and process lifetime management.
		/// </summary>
		public NavigationHelper NavigationHelper
		{
			get { return _navigationHelper; }
		}

		/// <summary>
		/// Gets the DefaultViewModel. This can be changed to a strongly typed view model.
		/// </summary>
		public ObservableDictionary DefaultViewModel
		{
			get { return _defaultViewModel; }
		}

		public HubPage()
		{
			InitializeComponent();
			_navigationHelper = new NavigationHelper(this);
			_navigationHelper.LoadState += NavigationHelper_LoadState;
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
			var sampleDataGroups = await SampleDataSource.GetGroupsAsync();
			DefaultViewModel["Groups"] = sampleDataGroups;

			// load our toggle switch states
			UserData.Provider = AppSettings.GetProviderO365() ? AuthProvider.O365 : AuthProvider.MicrosoftAccount;
			O365ToggleSwitch.IsOn = UserData.Provider == AuthProvider.O365;

			UserData.UseBeta = AppSettings.GetUseBeta();
			UseBetaToggleSwitch.IsOn = UserData.UseBeta;
		}

		/// <summary>
		/// Invoked when a HubSection header is clicked.
		/// </summary>
		/// <param name="sender">The Hub that contains the HubSection whose header was clicked.</param>
		/// <param name="e">Event data that describes how the click was initiated.</param>
		void Hub_SectionHeaderClick(object sender, HubSectionHeaderClickEventArgs e)
		{
			HubSection section = e.Section;
			var group = section.DataContext;
			Frame.Navigate(typeof(SectionPage), ((SampleDataGroup)group).UniqueId);
		}

		/// <summary>
		/// Invoked when an item within a section is clicked.
		/// </summary>
		/// <param name="sender">The GridView or ListView
		/// displaying the item clicked.</param>
		/// <param name="e">Event data that describes the item clicked.</param>
		void ItemView_ItemClick(object sender, ItemClickEventArgs e)
		{
			// Navigate to the appropriate destination page, configuring the new page
			// by passing required information as a navigation parameter
			UserData.ItemId = ((SampleDataItem)e.ClickedItem).UniqueId;
			Frame.Navigate(typeof(ItemPage), UserData);
		}
		#region NavigationHelper registration

		/// <summary>
		/// The methods provided in this section are simply used to allow
		/// NavigationHelper to respond to the page's navigation methods.
		/// Page specific logic should be placed in event handlers for the  
		/// <see cref="Common.NavigationHelper.LoadState"/>
		/// and <see cref="Common.NavigationHelper.SaveState"/>.
		/// The navigation parameter is available in the LoadState method 
		/// in addition to page state preserved during an earlier session.
		/// </summary>
		protected override void OnNavigatedTo(NavigationEventArgs e)
		{
			_navigationHelper.OnNavigatedTo(e);
		}

		protected override void OnNavigatedFrom(NavigationEventArgs e)
		{
			_navigationHelper.OnNavigatedFrom(e);
		}

		#endregion

		private void GroupSection_ItemClick(object sender, ItemClickEventArgs e)
		{
			var groupId = ((SampleDataGroup)e.ClickedItem).UniqueId;
			Frame.Navigate(typeof (SectionPage), groupId);
		}


		private async void O365ToggleSwitch_Toggled(object sender, RoutedEventArgs e)
		{
			var toggleSwitch = (ToggleSwitch)sender;
			var priorProviderSelection = UserData.Provider;
			UserData.Provider = toggleSwitch.IsOn ? AuthProvider.O365 : AuthProvider.MicrosoftAccount;
			AppSettings.SetProviderO365(toggleSwitch.IsOn);

			if (UserData.Provider != priorProviderSelection)
			{
				await Auth.SignOut(UserData.Provider);
			}
		}

		private void UseBetaToggleSwitch_Toggled(object sender, RoutedEventArgs e)
		{
			var toggleSwitch = (ToggleSwitch) sender;
			UserData.UseBeta = toggleSwitch.IsOn;
			AppSettings.SetUseBeta(UserData.UseBeta);
		}
	}
}
