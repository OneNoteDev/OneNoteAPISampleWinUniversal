using System;
using Windows.ApplicationModel.Resources;
using Windows.Graphics.Display;
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
		private readonly NavigationHelper _navigationHelper;
		private HubContext _hubContext;
		private readonly ObservableDictionary _defaultViewModel = new ObservableDictionary();

		public HubPage()
		{
			InitializeComponent();

			// Hub is only supported in Portrait orientation
			DisplayInformation.AutoRotationPreferences = DisplayOrientations.Portrait;

			NavigationCacheMode = NavigationCacheMode.Required;

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

		/// <summary>
		/// Gets the view model for this <see cref="Page"/>.
		/// This can be changed to a strongly typed view model.
		/// </summary>
		public ObservableDictionary DefaultViewModel
		{
			get { return _defaultViewModel; }
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
			// TODO: Create an appropriate data model for your problem domain to replace the sample data
			var sampleDataGroups = await SampleDataSource.GetGroupsAsync();
			DefaultViewModel["Groups"] = sampleDataGroups;
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

		/// <summary>
		/// Shows the details of a clicked group in the <see cref="SectionPage"/>.
		/// </summary>
		/// <param name="sender">The source of the click event.</param>
		/// <param name="e">Details about the click event.</param>
		private void GroupSection_ItemClick(object sender, ItemClickEventArgs e)
		{
			var groupId = ((SampleDataGroup)e.ClickedItem).UniqueId;
			Frame.Navigate(typeof (SectionPage), groupId);
		}

		/// <summary>
		/// Shows the details of an item clicked on in the <see cref="ItemPage"/>
		/// </summary>
		/// <param name="sender">The source of the click event.</param>
		/// <param name="e">Defaults about the click event.</param>
		private void ItemView_ItemClick(object sender, ItemClickEventArgs e)
		{
			_hubContext.ItemId = ((SampleDataItem)e.ClickedItem).UniqueId;
			Frame.Navigate(typeof(ItemPage), _hubContext);
		}

		#region NavigationHelper registration

		/// <summary>
		/// The methods provided in this section are simply used to allow
		/// NavigationHelper to respond to the page's navigation methods.
		/// <para>
		/// Page specific logic should be placed in event handlers for the
		/// The navigation parameter is available in the LoadState method
		/// in addition to page state preserved during an earlier session.
		/// </para>
		/// </summary>
		/// <param name="e">Event data that describes how this page was reached.</param>
		protected override void OnNavigatedTo(NavigationEventArgs e)
		{
			_navigationHelper.OnNavigatedTo(e);
		}

		protected override void OnNavigatedFrom(NavigationEventArgs e)
		{
			_navigationHelper.OnNavigatedFrom(e);
		}

		#endregion
	}
}