using Aspose.Words;
using Aspose.Words.Replacing;
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
        void Save(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "DOCX Files (*.docx)|*.docx|All Files (*.*)|*.*";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                doc.Range.Replace("DateD", DateD.ToString(), new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("DateM", DateM, new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("DateY", DateY.ToString(), new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("Number", Number.ToString(), new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("OrganizationType", OrganizationType, new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("OrganizationName", OrganizationName, new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("WHOM", WHOM, new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("OKOPF", OKOPF, new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("OKFS", OKFS, new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("ORGN", ORGN, new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("INN", INN, new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("KPP", KPP, new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("OKVED", OKVED, new FindReplaceOptions(FindReplaceDirection.Forward));

                doc.Range.Replace("PHONE", PHONE, new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("EMAIL", EMAIL, new FindReplaceOptions(FindReplaceDirection.Forward));

                doc.Range.Replace("AddresMailIndex", AddresMailIndex, new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("AddresArea", AddresArea, new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("AddresCity", AddresCity, new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("AddresStreet", AddresStreet, new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("AddresHome", AddresHome, new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("ReportArea", ReportArea, new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("ReportCity", ReportCity, new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("GUSZN", GUSZN, new FindReplaceOptions(FindReplaceDirection.Forward));


                doc.Range.Replace("DateWorkerHomeTimeBegin", DateWorkerHomeTimeBegin.ToString("dd.mm.yyyy"), new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("DateWorkerHomeTimeEnd", DateWorkerHomeTimeEnd.ToString("dd.mm.yyyy"), new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("WorkerHomeTimeI", WorkerHomeTimeI.ToString(), new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("WorkerHomeTime", WorkerHomeTime.ToString(), new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("WorkerHomeI", WorkerHomeI.ToString(), new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("WorkerHome", WorkerHome.ToString(), new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("WorkerI", WorkerI.ToString(), new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("Worker", Worker.ToString(), new FindReplaceOptions(FindReplaceDirection.Forward));
                
                doc.Save(dialog.FileName, SaveFormat.Docx);
                
            }
        }
    }
}