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

namespace StudentAttendance.Controllers
{
    public class StudentController : Controller
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Attendance Tracker";
        public String spreadsheetId = "14-UKYNismFO3QP9H20eMbkvumnzwnApoC9vu2VnMjTQ"; // Spreadsheet ID
        // GET: Student
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ClassRegister()
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
            var dinner = new ClassRegisterViewModel();
            foreach (var s in spr.Sheets)
            {
                dinner.classes.Add(s.Properties.Title);
            }
            return View(dinner);
        }
    }
}