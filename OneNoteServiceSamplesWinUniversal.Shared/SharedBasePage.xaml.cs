using System;
using Windows.System;
using Windows.UI;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Media;
using OneNoteServiceSamplesWinUniversal.OneNoteApi;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

namespace OneNoteServiceSamplesWinUniversal
{
	public class HubContext
	{
		public string ItemId;
		public AuthProvider Provider;
	}
    /// <summary>
    /// An empty base page that contains shared elements (e.g. BottomAppBar) that all other pages can use.
    /// All other pages in this project derive from this shared base page and inherit shared UI elements like
    /// the bottom app bar.
    /// </summary>
    /// <remarks>
    /// This class was built using the concepts defined at "How to share an app bar across pages (XAML)"
    /// http://msdn.microsoft.com/en-us/library/windows/apps/xaml/jj150604.aspx
    /// </remarks>
    public partial class SharedBasePage
    {
	    protected static HubContext UserData = new HubContext();
        public SharedBasePage()
        {
            InitializeComponent();
            Loaded += SharedBasePage_Loaded;
        }

        private void SharedBasePage_Loaded(object sender, RoutedEventArgs e)
        {
            var commandBar = new CommandBar
            {
                Background = new SolidColorBrush(Colors.Purple)
            };
            // This bottom commandBar contains the following buttons
            var devCenterButton = new AppBarButton
            {
                Icon = new SymbolIcon(Symbol.Home),
                Label = "Home",
                Tag = "http://dev.onenote.com",
            };
            devCenterButton.Click += ButtonBase_OnClick;
            commandBar.PrimaryCommands.Add(devCenterButton);

            var devBlogButton = new AppBarButton
            {
                Icon = new SymbolIcon(Symbol.World),
                Label = "Blog",
                Tag = "http://go.microsoft.com/fwlink/?LinkId=390183",
            };
            devBlogButton.Click += ButtonBase_OnClick;
            commandBar.PrimaryCommands.Add(devBlogButton);

            var otherSampleButton = new AppBarButton
            {
                Icon = new SymbolIcon(Symbol.Link),
                Label = "Samples",
                Tag = "http://go.microsoft.com/fwlink/?LinkId=390178",
            };
            otherSampleButton.Click += ButtonBase_OnClick;
            commandBar.PrimaryCommands.Add(otherSampleButton);

            var channel9VideoButton = new AppBarButton
            {
                Icon = new SymbolIcon(Symbol.Video),
                Label = "Videos",
                Tag = "http://channel9.msdn.com/Series/OneNoteDev",
            };
            channel9VideoButton.Click += ButtonBase_OnClick;
            commandBar.PrimaryCommands.Add(channel9VideoButton);

            var stackOverflowButton = new AppBarButton
            {
                Icon = new SymbolIcon(Symbol.Link),
                Label = "Forum",
                Tag = "http://go.microsoft.com/fwlink/?LinkId=390182",
            };
            stackOverflowButton.Click += ButtonBase_OnClick;
            commandBar.PrimaryCommands.Add(stackOverflowButton);

            var feedbackButton = new AppBarButton
            {
                Icon = new SymbolIcon(Symbol.Link),
                Label = "Feedback",
                Tag = "http://onenote.uservoice.com/forums/245490-onenote-developer-feedback-suggestions",
            };
            feedbackButton.Click += ButtonBase_OnClick;
            commandBar.PrimaryCommands.Add(feedbackButton);

            var twitterButton = new AppBarButton
            {
                Icon = new SymbolIcon(Symbol.Account),
                Label = "Twitter",
                Tag = "http://go.microsoft.com/fwlink/?LinkId=392528",
            };
            twitterButton.Click += ButtonBase_OnClick;
            commandBar.PrimaryCommands.Add(twitterButton);

            BottomAppBar = commandBar;
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            var button = sender as AppBarButton;
            if (button != null && button.Tag != null)
            {
                Launcher.LaunchUriAsync(new Uri(button.Tag.ToString()));
            }
        }

    }
}
