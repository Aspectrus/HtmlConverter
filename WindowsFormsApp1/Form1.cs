using Html2Markdown;
using System;
using System.IO;
using System.Net;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          string[] lines = System.IO.File.ReadAllLines(@"urls.txt");

        foreach(string  url in lines)
            {
                WebClient client = new WebClient();
                string htmlContent = client.DownloadString(url);
                //string editedHtmlContent = htmlContent.Substring(htmlContent.IndexOf("<title>"));
                try
                {
                   // editedHtmlContent = editedHtmlContent.Substring(0, editedHtmlContent.LastIndexOf("<div class=\"hero"));
                  //  editedHtmlContent = editedHtmlContent.Substring(0, editedHtmlContent.Length - 32);

                }
                catch (Exception ex)
                {

                }
                //  string filename = editedHtmlContent.Substring(7, editedHtmlContent.IndexOf("</title>")-9);
                //  filename = RemoveIllegalChars(filename);
                string text = System.IO.File.ReadAllText(@"Course 40336-A First Look Clinic Windows 10 for IT Professionals - Learn  Microsoft Do");


                //   CreateFile(editedHtmlContent, filename);
                var converter = new Html2Markdown.Converter();
                var md = converter.Convert(htmlContent);

            }    



        }

        private void CreateFile(string text, string filename)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter($"{filename}"))
            {
                file.WriteLine(text);
            }
        }

        private string RemoveIllegalChars(string filename)
        {
           
            string invalid = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());

            foreach (char c in invalid)
            {
                filename = filename.Replace(c.ToString(), "");
            }

            return filename;
        }
    }
}
