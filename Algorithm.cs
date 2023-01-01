using Fizzler.Systems.HtmlAgilityPack;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json;
using RestSharp;
using System.Diagnostics;
using System.Net;
using System.Text.RegularExpressions;

namespace PriceScraper
{
    internal class Algorithm
    {
        static string _environmentPath = "";
        static string _sheetId = "";
        static SheetsService _gService = null;
        static void Main(string[] args)
        {
            //Init environment
            _environmentPath = Path.GetFullPath(Path.Combine(@$"{Environment.CurrentDirectory}", @"..\..\.."));
            _gService = ImportantMethods.InitializeGoogleSheet();
            _sheetId = "1S0kkoXtqjR8hylly7rKnOKR9CXBacHt-umtex31CzKA";
            var allPhoneJsonSellMyMobile = GetPhoneJsonSellMyMobile();

            //algorithm
            while (true)
            {
                try
                {
                    var sheetObjs = GetSheetObjs();
                    for (int i = 0; i < sheetObjs.Count; i++)
                    {
                        try
                        {
                            if (sheetObjs[i].ProductName.Length < 1 || sheetObjs[i].ProductName.Contains(" ") == false) continue;

                            if (sheetObjs[i].SheetName == "Sell Pricing to scrape")
                            {
                                foreach (var phone in allPhoneJsonSellMyMobile)
                                {
                                    string name = "", url = "";

                                    try { name = ImportantMethods.FormatStringProperly(phone["name"]); } catch (Exception) { }
                                    try { url = "https://www.sellmymobile.com" + ImportantMethods.FormatStringProperly(phone["resultsUrl"]); } catch (Exception) { }

                                    if (name == sheetObjs[i].ProductName)
                                    {
                                        var urlResponseContent = GetResponse(linkToVisit: url, headerType: "sellmymobile", isPostRequest: false, body: "", useProxy: false);
                                        var document = ImportantMethods.GetHtmlNode(urlResponseContent);
                                        var scriptContent = document.QuerySelectorAll("script[type=\"text/javascript\"]").Where(x => x.InnerHtml.Contains("window.serverViewModel")).First().InnerHtml.Replace("window.serverViewModel = ", "").Replace("};", "}");
                                        var dealsJson = JsonConvert.DeserializeObject<dynamic>(scriptContent)["deals"];
                                        var allQuotes = new List<QuoteObj>();
                                        foreach (var deal in dealsJson)
                                        {
                                            string condition = "", quote = "";

                                            try { condition = ImportantMethods.FormatStringProperly(deal["condition"]); } catch (Exception) { }
                                            try { quote = ImportantMethods.FormatStringProperly(deal["quote"]); } catch (Exception) { }

                                            allQuotes.Add(new QuoteObj()
                                            {
                                                Condition = condition,
                                                Quote = quote
                                            });
                                        }

                                        sheetObjs[i].NewQuote = allQuotes.Where(x => x.Condition == "New").OrderByDescending(x => double.Parse(x.Quote)).First().Quote;
                                        sheetObjs[i].WorkingQuote = allQuotes.Where(x => x.Condition == "Working").OrderByDescending(x => double.Parse(x.Quote)).First().Quote;
                                        sheetObjs[i].BrokenQuote = allQuotes.Where(x => x.Condition == "Broken").OrderByDescending(x => double.Parse(x.Quote)).First().Quote;
                                        break;
                                    }
                                }
                            }
                            else if (sheetObjs[i].SheetName == "Ecom Pricing to scrape")
                            {
                                var responseContent = GetResponse(linkToVisit: $@"https://www.backmarket.com/en-us/search?q=", headerType: "backmarket", isPostRequest: false, body: "", useProxy: false);
                                var document = ImportantMethods.GetHtmlNode(responseContent);
                                var allItems = document.QuerySelectorAll("div[class=\"flex flex-row h-full items-center md:flex-col md:items-start\"]");
                                foreach (var item in allItems)
                                {
                                    string option = "", price = "", name = "";

                                    try { option = Regex.Match(item.QuerySelector("span[class=\"body-2-light duration-200 line-clamp-1 normal-case overflow-ellipsis overflow-hidden text-black transition-all\"]").InnerText, @".+?(?= - )").Value; } catch (Exception) { }
                                    try { price = ImportantMethods.FormatStringProperly(item.QuerySelector("span[class=\"body-2-bold text-black\"]").InnerText); } catch (Exception) { }
                                    try { name = ImportantMethods.FormatStringProperly(item.QuerySelector("h2[class=\"body-1-bold duration-200 line-clamp-1 md:mb-1 md:mt-0 mt-1 normal-case overflow-ellipsis overflow-hidden text-black transition-all\"]").InnerText).Split("\n").First(); } catch (Exception) { }

                                    if (option == sheetObjs[i].Option)
                                    {
                                        sheetObjs[i].Price = price;
                                        break;
                                    }
                                }
                            }
                        }
                        catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                    }
                    for (int i = 0, r = 1; i < sheetObjs.Count; i++, r++)
                    {
                        try
                        {
                            if (sheetObjs[i].ProductName.Length < 1 || sheetObjs[i].ProductName.Contains(" ") == false) continue;

                            if (sheetObjs[i].SheetName == "Sell Pricing to scrape")
                            {
                                ImportantMethods.UpdateSpreadsheet($"{sheetObjs[i].SheetName}!D{r}:D{r}", new List<object>() { sheetObjs[i].NewQuote }, _gService, _sheetId);
                                ImportantMethods.UpdateSpreadsheet($"{sheetObjs[i].SheetName}!E{r}:E{r}", new List<object>() { sheetObjs[i].WorkingQuote }, _gService, _sheetId);
                                ImportantMethods.UpdateSpreadsheet($"{sheetObjs[i].SheetName}!G{r}:G{r}", new List<object>() { sheetObjs[i].BrokenQuote }, _gService, _sheetId);
                            }
                            else if (sheetObjs[i].SheetName == "Ecom Pricing to scrape")
                            {
                                ImportantMethods.UpdateSpreadsheet($"{sheetObjs[i].SheetName}!G{r}:G{r}", new List<object>() { sheetObjs[i].Price }, _gService, _sheetId);
                            }
                            Thread.Sleep(TimeSpan.FromSeconds(2));
                        }
                        catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                    }
                    Thread.Sleep(TimeSpan.FromMinutes(0.1));
                }
                catch (Exception ex) { }
            }
        }
        private static dynamic GetPhoneJsonSellMyMobile()
        {
            return JsonConvert.DeserializeObject<dynamic>(GetResponse(linkToVisit: $@"https://www.sellmymobile.com/ajax/search/phones/", headerType: "sellmymobile", isPostRequest: false, body: "", useProxy: false));
        }
        public static string GetResponse(string linkToVisit, string headerType = "", bool isPostRequest = false, string body = "", bool useProxy = false)
        {
            var responseContent = "";
            for (int i = 0; i < 50 && responseContent.Length == 0 || responseContent.Contains("HTTP Error 400"); i++)
            {
                var client = new RestClient(linkToVisit)
                {
                    UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36",
                    //Proxy = new WebProxy()
                    //{
                    //    Address = new Uri("http://global.rotating.proxyrack.net:9000"),
                    //    Credentials = new NetworkCredential("myrate", "fd1544-6fd31d-562dd8-cab3de-e7b041")
                    //}

                };
                if (useProxy)
                {
                    client.Proxy = new WebProxy()
                    {
                        Address = new Uri($"http://user-suheyl:test123@gate.smartproxy.com:7000"),
                        Credentials = new NetworkCredential("suheyl", "test123")
                    };
                }
                var request = GetAppropriateHeader(headerType, isPostRequest, body);

                var response = client.ExecuteAsync(request);

                if (!response.Wait(300000)) { continue; };
                responseContent = response.Result.Content;
                if (responseContent.Length == 0 || responseContent.Contains("HTTP Error 400")) { continue; }
            }

            return responseContent;
        }
        [DebuggerStepThrough]
        private static RestRequest GetAppropriateHeader(string headerType, bool isPostRequest, string body)
        {
            var request = new RestRequest();
            IEnumerable<FileInfo> headerFile = null;

            if (isPostRequest) request = new RestRequest(Method.POST);

            headerFile = new DirectoryInfo($@"{_environmentPath}\Misc\ApiCookies").GetFiles().Where(x => x.Name == $"{headerType}.txt");
            var fileContent = File.ReadAllLines(headerFile.First().FullName);
            foreach (var line in fileContent)
            {
                try
                {
                    var headerName = Regex.Match(line, @".+?(?=\:)").Value.Trim();
                    var headerValue = Regex.Match(line, @"(?<=:).*").Value.Trim();

                    //if (headerName.ToLower().Contains("accept-encoding")) { headerValue = "application/json"; }

                    if (headerName.ToLower().Contains("user-agent")) { }
                    if (headerName.ToLower().Contains("content-length")) { }
                    else if (headerName.ToLower().Contains("host")) { }
                    else if (headerName.Contains("{") || headerValue.Contains("{")) { }
                    else request.AddHeader(headerName, headerValue);
                }
                catch (Exception) { }
            }
            if (body.Length > 0)
            {
                if (headerType == "services2")
                    request.AddParameter("application/x-www-form-urlencoded; charset=UTF-8", body, ParameterType.RequestBody);
                else
                    request.AddParameter("multipart/form-data; boundary=----WebKitFormBoundarywAjxGVrKcQAsXeO6", body, ParameterType.RequestBody);
            }

            return request;
        }
        private static List<SheetObj> GetSheetObjs()
        {
            var sheetObjs = new List<SheetObj>();
            var sheet1Data = ImportantMethods.ReadSpreadsheetEntries("Sell Pricing to scrape!C1:C9999", _gService, _sheetId);
            //var sheet2Data = ImportantMethods.ReadSpreadsheetEntries("Ecom Pricing to scrape!B1:C9999", _gService, _sheetId);
            sheetObjs.AddRange(sheet1Data.Select(x => new SheetObj() { SheetName = "Sell Pricing to scrape", ProductName = x.Count() > 0 ? x[0].ToString() : "" }));
            //sheetObjs.AddRange(sheet2Data.Select(x => new SheetObj() { SheetName = "Ecom Pricing to scrape", ProductName = x.Count() > 0 ? x[0].ToString() : "", Option = x.Count() > 0 ? x[1].ToString() : "" }));
            return sheetObjs;
        }
    }
    class QuoteObj
    {
        public string Quote { get; set; }
        public string Condition { get; set; }
    }
    class SheetObj
    {
        public string ProductName { get; set; }
        public string SheetName { get; set; }
        public string NewQuote { get; set; }
        public string WorkingQuote { get; set; }
        public string BrokenQuote { get; set; }
        public string Option { get; set; }
        public string Price { get; set; }
    }
}