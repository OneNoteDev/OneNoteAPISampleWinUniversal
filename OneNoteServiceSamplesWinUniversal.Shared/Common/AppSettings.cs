using System;
using System.Collections.Generic;
using System.Text;
using Windows.Storage;

namespace OneNoteServiceSamplesWinUniversal.Common
{
    public static class AppSettings
    {
		public static bool GetProviderO365()
		{
			object result = false;
			ApplicationData.Current.LocalSettings.Values.TryGetValue("ProviderO365", out result);
			return (bool)(result ?? false);
		}
		public static void SetProviderO365(bool isO365)
		{
			ApplicationData.Current.LocalSettings.Values["ProviderO365"] = isO365;
		}

		public static bool GetUseBeta()
		{
			object result = false;
			ApplicationData.Current.LocalSettings.Values.TryGetValue("UseBeta", out result);
			return (bool)(result ?? false); // default to true
		}
		public static void SetUseBeta(bool useBeta)
		{
			ApplicationData.Current.LocalSettings.Values["UseBeta"] = useBeta;
		}
	}
}
