using Aspose.Words;
using ExamsCDSP.Properties;
using System.Collections;
using System.Globalization;
using System.Reflection;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace ExamsCDSP
{
    public partial class Form1 : Form
    {
        static CultureInfo ci = CultureInfo.CurrentCulture;
        static DateTime Now = DateTime.Now;
        Document doc;
        int DateD = Now.Day;
        string DateM = ci.DateTimeFormat.GetMonthName(Now.Month);
        int DateY = Now.Year;
        int Number = 0;
        string OrganizationType = "";
        string OrganizationName = "";
        string WHOM = "";
        string OKOPF = "";
        string OKFS = "";
        string ORGN = "";
        string INN = "";
        string KPP = "";
        string OKVED = "";
        string PHONE = "";
        string EMAIL = "";
        string AddresMailIndex = "";
        string AddresArea = "";
        string AddresCity = "";
        string AddresStreet = "";
        string AddresHome = "";
        string ReportArea = "";
        string ReportCity = "";
        string GUSZN = "";


        int Worker = 0;
        int WorkerI = 0;
        int WorkerHome = 0;
        int WorkerHomeI = 0;
        int WorkerHomeTime = 0;
        int WorkerHomeTimeI = 0;
        DateTime DateWorkerHomeTimeBegin = Now;
        DateTime DateWorkerHomeTimeEnd = Now;




        public Form1()
        {
            doc = new Document(new MemoryStream(Resources.report));
            InitializeComponent();
        }
        void FileUpdate()
        {

        }
    }
}