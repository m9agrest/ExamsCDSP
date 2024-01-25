using Aspose.Words;
using ExamsCDSP.Properties;
using System.Collections;
using System.Reflection;

namespace ExamsCDSP
{
    public partial class Form1 : Form
    {
        Document doc;
        public Form1()
        {
            MemoryStream stream = new MemoryStream(Resources.report);
            doc = new Document(stream);
            InitializeComponent();
        }
    }
}