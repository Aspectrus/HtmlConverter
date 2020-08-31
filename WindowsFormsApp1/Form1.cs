using Html2Markdown;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        List<String> hrefslist = new List<String>();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          string[] lines = System.IO.File.ReadAllLines(@"urls.txt");
            int i = 0;
        foreach(string  url in lines)
            {
                WebClient client = new WebClient();
                string htmlContent = client.DownloadString(url);
                string title = TakeCourseTitle(htmlContent);
                string editedHtmlContent = "";
                try
                {
                    var contentStartIndex = htmlContent.IndexOf("<!-- <content> -->");
                    editedHtmlContent = htmlContent.Substring(contentStartIndex, htmlContent.IndexOf("<!-- </content> -->") - contentStartIndex);
                    editedHtmlContent = editedHtmlContent.Substring("<!-- <content> -->".Length);
                    if (editedHtmlContent.Contains("Chinese (Simplified)")) continue;
                    if (editedHtmlContent.Contains("Chinese (Traditional)")) continue;
                    if (editedHtmlContent.Contains("Japanese, English")) continue;

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }
                string filename = RemoveIllegalChars(title);


                var converter = new Html2Markdown.Converter();

                
                try
                {
                    var md = converter.Convert(editedHtmlContent);
                    md = StripHTML(md);
                    Console.WriteLine(filename + " " + i);
                    CreateFile(md, filename);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }
                i++;
            }

            MessageBox.Show("Finished");

        }

        private void CreateFile(string text, string filename)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter($"{filename}.md"))
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
            filename = filename.Replace("Â", "").Replace("®", "").Replace("Ã", "").Replace("©", "").Replace("â", "").Replace("€", "").Replace("™", "");

            return filename;
        }

        public static string StripHTML(string input)
        {

            input = Regex.Replace(input, "<[^>]*>", String.Empty).Replace("Show more", "").Replace("See all instructor-led courses", "");
            //input = Regex.Replace(input, @"\((.*?)spx\)", String.Empty);
            input = Regex.Replace(input, @"\[[^>]*spx\)|\![^>]*svg\)|\[[^>]*browse/\)", String.Empty);

            return Regex.Replace(input, @"# /  +/g\n|# \n", "#");

        }


        public string TakeCourseTitle(string htmlContent)
        {
            var titleStartIndex = htmlContent.IndexOf("<title>");
            string title = htmlContent.Substring(titleStartIndex, htmlContent.IndexOf("</title>") - titleStartIndex);
            return title.Substring("<title>".Length, title.Length - " -Learn | Microsoft Docs".Length - "<title>".Length);
        }

        public void GetAllHrefs(String uri)
        {



            var web = new HtmlWeb();
            web.BrowserTimeout = TimeSpan.FromSeconds(30);

            var doc = web.LoadFromBrowser(uri, o =>
            {
                var webBrowser = (WebBrowser)o;

                // Wait until the list shows up
                return webBrowser.Document.Body.InnerHtml.Contains("card-content-title");
            });

            // Show results
            foreach (var title in doc.DocumentNode.SelectNodes(".//a[@class='card-content-title']"))
            {
                hrefslist.Add(title.GetAttributeValue("href", String.Empty));
            }
            Console.ReadLine();


        }

        void CreateUrlFile(String filename)
        {
            TextWriter tw = new StreamWriter(filename);
            foreach (String href in hrefslist)
                tw.WriteLine("https://docs.microsoft.com/en-us" + href);

            tw.Close();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            bool pagesNotEmpty = true;
            var linkSkip = 0;

            while (pagesNotEmpty)
            {
                try
                {
                    GetAllHrefs("https://docs.microsoft.com/en-us/learn/certifications/courses/browse/?locales=en&skip=" + linkSkip.ToString());

                    linkSkip += 30;

                }
                catch (Exception)
                {
                    pagesNotEmpty = false;
                }

            }
            CreateUrlFile("urls.txt");
            MessageBox.Show("Finished");

        }
    }
}
