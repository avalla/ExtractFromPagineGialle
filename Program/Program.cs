using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using HtmlAgilityPack;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using ScrapySharp.Extensions;
using System.Text.RegularExpressions;

namespace ExtractFromPagineGialle
{
    class Program
    {
        private readonly static String OutputDir = System.Configuration.ConfigurationManager.AppSettings["OutputDir"] ?? ".";
        private readonly static String Proxy = System.Configuration.ConfigurationManager.AppSettings["Proxy"] ?? "false";
        private readonly static String Login = System.Configuration.ConfigurationManager.AppSettings["Login"];
        private readonly static String Pwd = System.Configuration.ConfigurationManager.AppSettings["Password"];
        private readonly static String Dominio = System.Configuration.ConfigurationManager.AppSettings["Dominio"];
        private readonly static String ProxyAddress = System.Configuration.ConfigurationManager.AppSettings["ProxyAddress"];
        private readonly static Int32 ProxyPort = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["ProxyPort"] ?? "8080");
        static void Main(string[] args)
        {
            Console.WriteLine("> Cosa stai cercando?");
            var cosa = Console.ReadLine();
            cosa = HttpUtility.HtmlEncode(cosa ?? "commercialista");
            Console.WriteLine("> Dove? (Città / CAP)");
            var dove = Console.ReadLine();
            dove = HttpUtility.HtmlEncode(dove ?? "10100");
            if (String.IsNullOrWhiteSpace(cosa)) cosa = HttpUtility.HtmlEncode("commercialista");
            if (String.IsNullOrWhiteSpace(dove)) dove = HttpUtility.HtmlEncode("10100");
            var fileName = "estrazione_" + dove.Replace(" ", "_") + "_" + cosa.Replace(" ", "_") + ".xlsx";
            var url = String.Format("http://www.paginegialle.it/pgol/4-{0}/3-{1}", cosa, dove);
            Console.WriteLine("\nUrl {0}", url);
            var htmlDocument = EstraiDocument(url);
            #if DEBUG
                htmlDocument.Save(String.Format("{0}\\{1}-0.html",OutputDir, cosa));
            #endif
            var html = htmlDocument.DocumentNode;
            var foundItems = 0;
            var totalPages = 0;
            var spanTotal =
                html.CssSelect("div.content.list-main div.list-left h3.title_listing span.h-bold").FirstOrDefault();
            if (spanTotal != null)
                foundItems = Convert.ToInt32(spanTotal.InnerText);
            var spanPage =
                html.CssSelect("div.content.list-main div.list-left div.footer-listing div.pag-group p.mostra-foot span.bold-foot").LastOrDefault();
            if (spanPage != null)
                totalPages = Convert.ToInt32(spanPage.InnerText);
            Console.WriteLine("Trovati {0} elementi", foundItems);
            Console.WriteLine("Trovate {0} pagine", totalPages);
            var wait = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Wait"]);
            var commercianti = new List<Commerciante>();
            for (var i = 1; i <= totalPages; i++)
            {
                var urlToAnalyze = String.Format("http://www.paginegialle.it/pgol/4-{0}/3-{1}/p-{2}", cosa, dove, i);
                Console.WriteLine("\nProcesso la pagina {0}", i.ToString(CultureInfo.InvariantCulture));
                Console.WriteLine("\nUrl {0}", urlToAnalyze);
                var htmlDocPage = EstraiDocument(urlToAnalyze);
#if DEBUG
                htmlDocPage.Save(String.Format("{0}\\{1}-{2}.html", OutputDir, cosa, i));
#endif
                var htmlToAnalyze = htmlDocPage.DocumentNode;
                var list = htmlToAnalyze.CssSelect("div.content.list-main div.list-left div div.item.clearfix");
                Console.WriteLine("Trovati {0} elementi", list.Count().ToString(CultureInfo.InvariantCulture));
                var estrazione = EstraiInfo(list);
                commercianti.AddRange(estrazione);
                Console.WriteLine("Estratti {0} elementi", commercianti.Count().ToString(CultureInfo.InvariantCulture));
                System.Threading.Thread.Sleep(wait);
            }
            if (commercianti.Any())
                SalvaXls(commercianti, fileName, OutputDir);

        }

        private static MemoryStream GetPage(string url, bool useProxy)
        {
            var uri = new Uri(url);
            var wc = new WebClient();
            wc.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; .NET CLR 1.1.4322; .NET CLR 2.0.50727)");
            wc.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8");
            wc.Headers.Add(HttpRequestHeader.Referer, "http://www.google.com");
            if (useProxy) wc.Proxy  = new WebProxy(ProxyAddress, ProxyPort) { UseDefaultCredentials = true };
            //request.Credentials = new NetworkCredential(Login, Pwd, Dominio);
            return new MemoryStream(wc.DownloadData(uri));
        }

        private static HtmlDocument EstraiDocument(string url)
        {
            var data = GetPage(url, Proxy == "true");
            var htmlDoc = new HtmlDocument();
            htmlDoc.Load(data);
            return htmlDoc;
        }

        private static IEnumerable<Commerciante> EstraiInfo(IEnumerable<HtmlNode> list)
        {
            var result = new List<Commerciante>();
            foreach (var htmlNode in list)
            {
                var item = new Commerciante();
                var name = htmlNode.CssSelect("div.item_sx div.item_head div.org.fn h2.rgs a").FirstOrDefault();
                if (name != null) {
                    item.Nome = name.InnerText.Ripulisci();
                    item.LinkPagineGialle = name.Attributes["href"].Value;
                }
                var indirizzo = htmlNode.CssSelect("div.item_sx div.address").FirstOrDefault();
                if (indirizzo != null)
                {
                    var street = indirizzo.CssSelect("span.street-address").FirstOrDefault();
                    if (street != null)
                        item.Indirizzo = street.InnerText.Ripulisci();
                    var locality = indirizzo.CssSelect("div.locality").FirstOrDefault();
                    if (locality != null)
                    {
                        var rx = new Regex(@"^\d{5}");
                        var loc = locality.InnerText.Ripulisci();
                        var cap = rx.Match(loc).Value;
                        item.Cap = cap;
                        item.Citta = rx.Replace(loc, String.Empty).Trim();

                    }
                    var telefono = indirizzo.CssSelect("div.tel").FirstOrDefault();
                    if (telefono != null)
                    {
                        foreach (var contatto in telefono.ChildNodes)
                        {
                            if (contatto != null)
                                item.Contatti += contatto.InnerText.Ripulisci();
                        }
                    }
                }

                var link = htmlNode.CssSelect("div.item_sx div.link a").FirstOrDefault();
                if (link != null)
                    item.Link = link.Attributes.Where(x => x.Name == "href").Select(x => x.Value).FirstOrDefault();
                var desc = htmlNode.CssSelect("div.item_sx div.text p.abstract").FirstOrDefault();
                if (desc != null)
                    item.Descrizione = desc.InnerText.Ripulisci();
                result.Add(item);
            }
            return result;
        }

        private static void SalvaXls(IEnumerable<Commerciante> list, string fileName, string outputDir)
        {
            var newFile = new FileInfo(outputDir + @"\" + fileName);
			if (newFile.Exists)
			{
				newFile.Delete();  
				newFile = new FileInfo(outputDir + @"\" + fileName);
			}
			using (var package = new ExcelPackage(newFile))
			{
                var worksheet = package.Workbook.Worksheets.Add("Estrazione");
                worksheet.Cells["A1"].LoadFromCollection(list, true, TableStyles.Light1);
                package.Save();
			}
            Console.WriteLine("\nSalvato file {0}",fileName);
        }
    }
}
