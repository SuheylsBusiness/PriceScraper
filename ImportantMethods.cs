using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Services;
using Newtonsoft.Json;
using System.Runtime.InteropServices;
using Path = System.IO.Path;
using System.Diagnostics;
using System.Web;
using HtmlAgilityPack;

namespace PriceScraper
{
    [DebuggerStepThrough]
    public class ImportantMethods
    {
        public static void UpdateSpreadsheet(string range, List<object> objects, SheetsService sheetService, string googleSheet_sheetID)
        {
            var valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>> { objects };
            var appendRequest = sheetService.Spreadsheets.Values.Update(valueRange, googleSheet_sheetID, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();
        }
        public static HtmlNode GetHtmlNode(string responseStr)
        {
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(responseStr);
            var document = htmlDoc.DocumentNode;
            return document;
        }
        public static string FormatStringProperly(object inputStr)
        {
            return HttpUtility.HtmlDecode(inputStr.ToString()).Trim();
        }
        public static void ClearSheet(string range, SheetsService sheetService, string googleSheet_sheetID)
        {
            // TODO: Assign values to desired properties of `requestBody`:
            Google.Apis.Sheets.v4.Data.ClearValuesRequest requestBody = new Google.Apis.Sheets.v4.Data.ClearValuesRequest();

            SpreadsheetsResource.ValuesResource.ClearRequest request = sheetService.Spreadsheets.Values.Clear(requestBody, googleSheet_sheetID, range);

            // To execute asynchronously in an async method, replace `request.Execute()` as shown:
            Google.Apis.Sheets.v4.Data.ClearValuesResponse response = request.Execute();

            // TODO: Change code below to process the `response` object:
            Console.WriteLine(JsonConvert.SerializeObject(response));
        }
        public static void SendEmail(string subject, string body, string destinationEmail, List<string> attachements)
        {
            string username = "suheylupworkbot@gmail.com", password = "stumynehzzollpzo";
            using (System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient("smtp.gmail.com", 587))
            {
                client.EnableSsl = true;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Credentials = new NetworkCredential(username, password);
                MailMessage msgObj = new MailMessage();
                msgObj.To.Add(destinationEmail);
                msgObj.From = new System.Net.Mail.MailAddress(username);
                msgObj.Subject = subject;
                msgObj.Body = body;
                msgObj.IsBodyHtml = true;
                for (int i = 0; i < attachements.Count; i++)
                    msgObj.Attachments.Add(new Attachment(attachements[i]));
                client.Send(msgObj);
            }
        }
        public static void SendEmail(string subject, string body, string destinationEmail)
        {
            string username = "suheylupworkbot@gmail.com", password = "stumynehzzollpzo";
            using (System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient("smtp.gmail.com", 587))
            {
                client.EnableSsl = true;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Credentials = new NetworkCredential(username, password);
                MailMessage msgObj = new MailMessage();
                msgObj.To.Add(destinationEmail);
                msgObj.From = new System.Net.Mail.MailAddress(username);
                msgObj.Subject = subject;
                msgObj.Body = body;
                msgObj.IsBodyHtml = true;
                client.Send(msgObj);
            }
        }
        public static void AppentIntoTopSpreadsheet2(List<string> objects, SheetsService sheetService, string googleSheet_sheetID, int sheetId)
        {
            InsertDimensionRequest insertRow = new InsertDimensionRequest();
            insertRow.Range = new DimensionRange()
            {
                SheetId = sheetId,
                Dimension = "ROWS",
                StartIndex = 1,
                EndIndex = 2
            };

            PasteDataRequest data = new PasteDataRequest
            {
                Data = string.Join(";%_", objects),
                Delimiter = ";%_",
                Coordinate = new GridCoordinate
                {
                    ColumnIndex = 0,
                    RowIndex = 1,
                    SheetId = sheetId
                },
            };

            BatchUpdateSpreadsheetRequest r = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>
                {
                    new Request{ InsertDimension = insertRow },
                    new Request{ PasteData = data }
                }
            };

            BatchUpdateSpreadsheetResponse response1 = sheetService.Spreadsheets.BatchUpdate(r, googleSheet_sheetID).Execute();

        }
        public static void AppentIntoSpreadsheet(string range, List<object> objects, SheetsService sheetService, string googleSheet_sheetID)
        {
            var valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>> { objects };
            var appendRequest = sheetService.Spreadsheets.Values.Append(valueRange, googleSheet_sheetID, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();
        }
        public static void AppentIntoTopSpreadsheet(List<string> objects, SheetsService sheetService, string googleSheet_sheetID, int sheetId)
        {
            InsertDimensionRequest insertRow = new InsertDimensionRequest();
            insertRow.Range = new DimensionRange()
            {
                SheetId = sheetId,
                Dimension = "ROWS",
                StartIndex = 1,
                EndIndex = 2
            };

            PasteDataRequest data = new PasteDataRequest
            {
                Data = string.Join(";%_", objects),
                Delimiter = ";%_",
                Coordinate = new GridCoordinate
                {
                    ColumnIndex = 0,
                    RowIndex = 1,
                    SheetId = sheetId
                },
            };

            BatchUpdateSpreadsheetRequest r = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>
                {
                    new Request{ InsertDimension = insertRow },
                    new Request{ PasteData = data }
                }
            };

            BatchUpdateSpreadsheetResponse response1 = sheetService.Spreadsheets.BatchUpdate(r, googleSheet_sheetID).Execute();

        }
        public static IList<IList<object>> ReadSpreadsheetEntries(string range, SheetsService sheetService, string googleSheet_sheetID)
        {
            var request = sheetService.Spreadsheets.Values.Get(googleSheet_sheetID, range);
            var response = request.Execute();
            var values = response.Values;
            return values;
        }
        public static SheetsService InitializeGoogleSheet()
        {
            GoogleCredential credential;
            string[] googleSheet_scopes = { SheetsService.Scope.Spreadsheets };
            using (var stream = new FileStream("projectaffordany-1587302173157-d20f1c21561a.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(googleSheet_scopes);
            }

            SheetsService sheetService;
            sheetService = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer() { HttpClientInitializer = credential, ApplicationName = "Task", });
            return sheetService;
        }
        public static List<string> removeDuplicatesFromList(List<string> input)
        {
            try { input = input.GroupBy(x => x).Select(x => x.First()).ToList(); } catch (Exception) { }
            return input;
        }
    }
}
