using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using StudentAttendance.Models;
using System.Diagnostics;
using System.Net;

namespace StudentAttendance.Controllers
{
    public class HomeController : Controller
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Google Sheets API .NET Quickstart";
        StringBuilder signInCell = new StringBuilder(); // Cell to sign in user
        String day = DateTime.Today.ToString("d"); // Gets today's date
        // Find the date column and put into 'dateColPosition' to append to 'signInCell'
        string[] dateColPosition = new string[] { "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
        // Generate unique 5 digit key needed for login
        public static int professorKey = new Random().Next(10000, 99999);
        public String spreadsheetId = "14-UKYNismFO3QP9H20eMbkvumnzwnApoC9vu2VnMjTQ"; // Spreadsheet ID
        String stdntIDCol = "D2:D40"; // Column 'D' of Student IDs
        String dateRow = "F1:Z1"; // Row 1 of dates

        //
        // GET: /Home/
        // Startup method
        public ActionResult Index()
        {
            /****** Gets IP Address and puts into string localIP. 'Debug.WriteLine' displays it while in debug mode. ******/
            IPHostEntry host;
            string localIP = "?";
            host = Dns.GetHostEntry(Dns.GetHostName());
            foreach (IPAddress ip in host.AddressList)
            {
                if (ip.AddressFamily.ToString() == "InterNetwork")
                {
                    localIP = ip.ToString();
                    Debug.WriteLine(localIP);
                }
            }
            /*************************************************************************/
            ViewBag.Key = professorKey; // Assign key to ViewBag object and pass to the view (Index.cshtml)
            /***This shows student what key is. Temporary placement since clearly we don't want just anyone knowing what the key is. Just for testing purposes***/
            return View();
        }

        //
        // GET: /About/
        public ActionResult About()
        {
            return View();
        }

        //
        // GET: /Contact/
        public ActionResult Contact()
        {
            return View();
        }

        // Method 'Updated' called when user submits information from Index view. Pass model as parameter
        // to compare user input with list of student IDs & unique key
        [HttpPost]
        public ActionResult Updated(AttendanceViewModel ui)
        {
            /******************************** Credential Stuff ***************************************/
            UserCredential credential;
            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            /*****************************************************************************************/

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            /***********************************/

            // A GET request for list of student IDs
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, stdntIDCol);
            // A GET request for list of dates
            SpreadsheetsResource.ValuesResource.GetRequest request1 =
                    service.Spreadsheets.Values.Get(spreadsheetId, dateRow);

            // Requests from spreadsheets returns values of type 'list of lists'
            // Execute request for list of dates and store into 'dates' variable
            IList<IList<Object>> dates = request1.Execute().Values;
            // Execute request for list of student IDs and store into 'stdIDs' variable
            IList<IList<Object>> stdIDs = request.Execute().Values;

            signInCell.Append(searchDate(dates)); // Append column letter based on today's date (E for not found)
            signInCell.Append(searchStdntID(stdIDs, ui)); // Append row number based on student ID match (0 for not found)

            if (signInCell.ToString()[0] != 'E') // If date found, continue
            {
                if (signInCell.ToString()[1] != '0') // If student ID found, continue
                {
                    if (ui.userKey == professorKey) // If user entered correct key, continue
                    {
                        ValueRange valueRange = new ValueRange(); // ValueRange object to update cells
                        // valueRange.MajorDimension = "COLUMNS";          //"ROWS";//"COLUMNS";// May use later
                        var oblist = new List<object>() { "Y" }; // Add 'Y' to list
                        valueRange.Values = new List<IList<object>> { oblist }; // Add list of list to ValueRange object
                        // Prepare to update spreadsheet with valueRange
                        SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, signInCell.ToString());
                        // ??
                        update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                        update.Execute(); // Execute request to update
                        return View(); // Return 'Updated' view
                    }
                    else
                    {
                        // Send error message to view
                        TempData["Message"] = "Incorrect key.";
                        return RedirectToAction("Error");
                    }
                }
                else
                {
                    // Send error message to view
                    TempData["Message"] = "Student ID not found.";
                    return RedirectToAction("Error");
                }
            }
            else
            {
                // Send error message to view
                TempData["Message"] = "Cannot find date.";
                return RedirectToAction("Error");
            }
        }

        // Render 'Error' view if date, student ID, or key is incorrect
        public ActionResult Error()
        {
            ViewBag.Key = professorKey; // Again, pass key to view for testing purposes
            return View();
        }

        // Method to find today's date located in 'date' variable
        private string searchDate(IList<IList<Object>> dates)
        {
            int index = 0;
            foreach (var row in dates)
            {
                foreach (var column in row)
                {
                    if (day == (string)column)
                        return dateColPosition[index]; // Gets letter based on array index
                    index++;
                }
            }
            return "E"; // If date not found, return E
        }

        // Method to find student ID
        private int searchStdntID(IList<IList<Object>> stdids, AttendanceViewModel ui)
        {
            int index = 2;
            if (stdids != null && stdids.Count > 0)
            {
                foreach (var row in stdids)
                {
                    foreach (var column in row)
                    {
                        if (ui.studentID == (string)column) // If ID found return second part of 'signInCell'
                        {
                            return index;
                        }
                    }
                    index++;
                }
            }
            return 0; // If student ID not found, return 0
        }

    }
}
