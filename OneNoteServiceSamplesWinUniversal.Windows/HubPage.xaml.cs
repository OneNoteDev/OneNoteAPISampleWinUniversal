using System;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Navigation;
using OneNoteServiceSamplesWinUniversal.Common;
using OneNoteServiceSamplesWinUniversal.Data;

// The Universal Hub Application project template is documented at http://go.microsoft.com/fwlink/?LinkID=391955
namespace OneNoteServiceSamplesWinUniversal
{
	/// <summary>
	/// A page that displays a grouped collection of items.
	/// </summary>
	public sealed partial class HubPage : SharedBasePage
	{
		private NavigationHelper navigationHelper;
		private ObservableDictionary defaultViewModel = new ObservableDictionary();

		/// <summary>
		/// Gets the NavigationHelper used to aid in navigation and process lifetime management.
		/// </summary>
		public NavigationHelper NavigationHelper
		{
			get { return navigationHelper; }
		}

		/// <summary>
		/// Gets the DefaultViewModel. This can be changed to a strongly typed view model.
		/// </summary>
		public ObservableDictionary DefaultViewModel
		{
			get { return defaultViewModel; }
		}

		public HubPage()
		{
			InitializeComponent();
			navigationHelper = new NavigationHelper(this);
			navigationHelper.LoadState += NavigationHelper_LoadState;
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
			var itemId = ((SampleDataItem)e.ClickedItem).UniqueId;
			Frame.Navigate(typeof(ItemPage), itemId);
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
			navigationHelper.OnNavigatedTo(e);
		}

		protected override void OnNavigatedFrom(NavigationEventArgs e)
		{
			navigationHelper.OnNavigatedFrom(e);
		}

		#endregion

		private void GroupSection_ItemClick(object sender, ItemClickEventArgs e)
		{
			var groupId = ((SampleDataGroup)e.ClickedItem).UniqueId;
			Frame.Navigate(typeof (SectionPage), groupId);
		}
	}
}
