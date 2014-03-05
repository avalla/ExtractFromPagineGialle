using System;
using System.Web;

namespace ExtractFromPagineGialle
{
    static class StringExtensionMethod
    {
        public static string Ripulisci(this string s)
        {
            s = s.Replace(Environment.NewLine, String.Empty).Replace("\n", String.Empty).Replace("\r\n", String.Empty).Trim(); 
            return HttpUtility.HtmlDecode(s);
        }
    }
}
