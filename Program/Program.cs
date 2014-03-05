using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using HtmlAgilityPack;
using OfficeOpenXml;
using ScrapySharp.Extensions;
using System.Text.RegularExpressions;

namespace ExtractFromPagineGialle
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Cosa?");
            var cosa = Console.ReadLine();
            Console.WriteLine("Dove?");
            var dove = Console.ReadLine();
            if (String.IsNullOrEmpty(cosa))
                cosa = "commercialista";
            if (String.IsNullOrEmpty(dove))
                dove = "10100";
            var fileName = "estrazione_" + dove.Replace(" ", "_") + "_" + cosa.Replace(" ", "_") + ".xlsx";
            cosa = HttpUtility.HtmlEncode(cosa);
            dove = HttpUtility.HtmlEncode(dove);
            var url = "http://www.paginegialle.it/pgol/4-" + cosa + "/3-" + dove + "/p-1?mr=50";
            var html = EstraiElementi(url);

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
            Console.WriteLine("Trovati {0} elementi", foundItems.ToString(CultureInfo.InvariantCulture));
            Console.WriteLine("Trovate {0} pagine", totalPages.ToString(CultureInfo.InvariantCulture));
            var wait = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Wait"]);
            var commercianti = new List<Commerciante>();
            for (var i = 1; i <= totalPages; i++)
            {
                Console.WriteLine("\nProcesso la pagina {0}", i.ToString(CultureInfo.InvariantCulture));
                var urlToAnalyze =  "http://www.paginegialle.it/pgol/4-"  + cosa + "/3-" + dove + "/p-" + i.ToString(CultureInfo.InvariantCulture) + "?mr=50";
                var htmlToAnalyze = EstraiElementi(urlToAnalyze);
                var list = htmlToAnalyze.CssSelect("div.content.list-main div.list-left div div.item.clearfix");
                Console.WriteLine("Trovati {0} elementi", list.Count().ToString(CultureInfo.InvariantCulture));
                var estrazione = EstraiInfo(list);
                commercianti.AddRange(estrazione);
                Console.WriteLine("Estratti {0} elementi", commercianti.Count().ToString(CultureInfo.InvariantCulture));
                System.Threading.Thread.Sleep(wait);
            }
            if (commercianti.Any())
                SalvaXls(commercianti, fileName);

        }
        private static HtmlNode EstraiElementi(string url)
        {
            var request = WebRequest.Create(url) as HttpWebRequest;
            var proxy = Convert.ToBoolean(System.Configuration.ConfigurationManager.AppSettings["Proxy"]);
            if (proxy) {
                // PROXY
                var login = System.Configuration.ConfigurationManager.AppSettings["Login"];
                var pwd = System.Configuration.ConfigurationManager.AppSettings["Password"];
                var dominio = System.Configuration.ConfigurationManager.AppSettings["Dominio"];
                var proxyAddress = System.Configuration.ConfigurationManager.AppSettings["ProxyAddress"];
                var proxyPort = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["ProxyPort"]);
                var cred = new NetworkCredential(login, pwd, dominio);
                var wp = new WebProxy(proxyAddress, proxyPort) { UseDefaultCredentials = true };
                request.Credentials = cred;
                request.Proxy = wp;
            }
            request.CookieContainer = new CookieContainer();
            request.UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; .NET CLR 1.1.4322; .NET CLR 2.0.50727)";
            request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
            var response = request.GetResponse() as HttpWebResponse;
            if (response == null)
                throw new Exception("Response null");
            if (response.StatusCode != HttpStatusCode.OK)
                throw new Exception("Risposta != 200");
            var htmlDoc = new HtmlDocument();
            htmlDoc.Load(response.GetResponseStream());
            response.Close();
            return htmlDoc.DocumentNode;
        }

        private static IEnumerable<Commerciante> EstraiInfo(IEnumerable<HtmlNode> list)
        {
            var result = new List<Commerciante>();
            foreach (var htmlNode in list)
            {
                var item = new Commerciante();
                var name = htmlNode.CssSelect("div.item_sx div.item_head div.org.fn h2.rgs a").FirstOrDefault();
                if (name != null)
                    //item.Nome = name.Attributes.Where(x => x.Name == "title").Select(x => x.Value).FirstOrDefault();
                    item.Nome = name.InnerText.Ripulisci();
                var indirizzo = htmlNode.CssSelect("div.item_sx div.address").FirstOrDefault();
                if (indirizzo != null)
                {
                    var street = indirizzo.CssSelect("span.street-address").FirstOrDefault();
                    if (street != null)
                        item.Indirizzo = street.InnerText.Ripulisci();
                    var locality = indirizzo.CssSelect("span.locality").FirstOrDefault();
                    if (locality != null)
                    {
                        var rx = new Regex(@"^\d{5}");
                        var loc = locality.InnerText.Ripulisci();
                        var cap = rx.Match(loc).Value;
                        item.Cap = cap;
                        item.Citta = rx.Replace(loc, String.Empty).Trim();
                    }
                }
                var telefono = htmlNode.CssSelect("div.item_sx div.address div div.tel").FirstOrDefault();
                if (telefono != null) {
                    if (telefono.InnerText.Contains("tel"))
                    {   
                        var telNumber = telefono.CssSelect("span.value").FirstOrDefault();
                        item.Telefono = telNumber.InnerText.Ripulisci();
                    }
                    if (telefono.InnerText.Contains("fax"))
                    {
                        var faxNumber = telefono.CssSelect("span.value").LastOrDefault();
                        item.Fax = faxNumber.InnerText.Ripulisci();
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

        private static void SalvaXls(IEnumerable<Commerciante> list, string fileName)
        {
            var outputDir = System.Configuration.ConfigurationManager.AppSettings["OutputDir"];
            var newFile = new FileInfo(outputDir + @"\" + fileName);
			if (newFile.Exists)
			{
				newFile.Delete();  
				newFile = new FileInfo(outputDir + @"\" + fileName);
			}
			using (var package = new ExcelPackage(newFile))
			{
                var worksheet = package.Workbook.Worksheets.Add("Estrazione");
                worksheet.Cells["A1"].LoadFromCollection(list);
                package.Save();
			}
            Console.WriteLine("\nSalvato file {0}",fileName);
        }
    }
}
