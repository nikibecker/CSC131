using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using StudentAttendance.Models;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using System.Diagnostics;

namespace StudentAttendance.Controllers
{
    public class ProfessorController : Controller
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Attendance Tracker";
        public String spreadsheetId = "14-UKYNismFO3QP9H20eMbkvumnzwnApoC9vu2VnMjTQ"; // Spreadsheet ID
        public static ProfessorViewModel profObj;
        ApplicationDbContext context = new ApplicationDbContext();
        // GET: Professor
        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult GenerateKey()
        {
            profObj = new ProfessorViewModel(new Random().Next(10000, 99999));
            return PartialView("KeyPartial", profObj);
        }

        [HttpGet]
        public ActionResult GenerateFakeKey()
        {
            return PartialView("FakePartial");
        }

        [HttpGet]
        public ActionResult CreateClassForm()
        {
            return PartialView("CreateClassPartial");
        }

        [HttpPost]
        public ActionResult AddStudent(ClassRegisterViewModel course)
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
            /****************************************************************************************/
            Spreadsheet spr = new Spreadsheet();
            spr = new SpreadsheetsResource(service).Get(spreadsheetId).Execute();

            String theclass = "A2:D40";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, theclass);
            IList<IList<Object>> classvalues = request.Execute().Values;

            ValueRange valueRange = new ValueRange(); // ValueRange object to update cells
                                                      // valueRange.MajorDimension = "COLUMNS";          //"ROWS";//"COLUMNS";// May use later
            var oblist = new List<object>() { course.lname, course.fname, User.Identity.Name, course.sid }; // Add 'Y' to list
            valueRange.Values = new List<IList<object>> { oblist }; // Add list of list to ValueRange object
                                                                    // Prepare to update spreadsheet with valueRange
            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, course.sheet + "A1:D1");
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            update.Execute(); // Execute request to update
            
            return PartialView("ClassAdded");
        }

        [HttpPost]
        public ActionResult CreateClass(CreateClassViewModel model)
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
            /****************************************************************************************/

            string sheetName = model.basedGodClassID + " - " + model.basedGodSectionNum;
            var addSheetRequest = new AddSheetRequest();
            addSheetRequest.Properties = new SheetProperties();
            addSheetRequest.Properties.Title = sheetName;
            BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateSpreadsheetRequest.Requests = new List<Request>();
            batchUpdateSpreadsheetRequest.Requests.Add(new Request
            {
                AddSheet = addSheetRequest
            });
            var batchUpdateRequest =
                service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, spreadsheetId);

            batchUpdateRequest.Execute();

            return RedirectToAction("Index");
        }
    }
}