using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace StudentAttendance.Models
{
    public class AttendanceViewModel
    {
        public string studentID { get; set; }
        public int userKey { get; set; }
    }

    public class ProfessorViewModel
    {
        public int uniqueKey { get; set; }

        public ProfessorViewModel(int key)
        {
            uniqueKey = key;
        }
    }

    public class CreateClassViewModel
    {
        public string basedGodClassID { get; set; }
        public int basedGodSectionNum { get; set; }
    }

    public class ClassRegisterViewModel
    {
        public IList<string> classes { get; set; }
        public Sheet sheet { get; set; }
        public string fname { get; set; }
        public string lname { get; set; }
        public string sid { get; set; }

        public ClassRegisterViewModel()
        {
            classes = new List<string>();
        }
    }
}