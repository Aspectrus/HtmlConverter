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



        ExcelFileGenerator excelFileGenerator = new ExcelFileGenerator();
        HashSet<string> hrefslist  = new HashSet<string>();
        IDictionary<String, String> prodDictionary = new Dictionary<string, string>();
     
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            UrlGeneration();

            string[] lines = System.IO.File.ReadAllLines(@"urls.txt");
          int a = 0, b = 0;
          foreach (string  url in lines)
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
                    if (!editedHtmlContent.Contains("<li>English</li>")) continue;
                    if (editedHtmlContent.Contains("Chinese (Simplified)")) continue;
                    if (editedHtmlContent.Contains("Chinese (Traditional)")) continue;
                    if (editedHtmlContent.Contains("Japanese, English</li>")) continue;
                    if (editedHtmlContent.Contains("German, English</li>")) continue;
                    if (editedHtmlContent.Contains("French, English</li>")) continue;
                    if (editedHtmlContent.Contains("Course retirement date:")) continue;
                    if (editedHtmlContent.Contains("Portuguese(Brazil), English")) continue;
                    if (editedHtmlContent.Contains("Spanish, English</li>>")) continue;
                 


                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }
                string filename = RemoveIllegalChars(title) + ".md";


                excelFileGenerator.CreateCourseListInfoFromHtml(editedHtmlContent, title,filename,url, prodDictionary);
                var converter = new Html2Markdown.Converter();


                

                try
                {
                    var md = converter.Convert(editedHtmlContent);
                    b = md.IndexOf("English");
                    a =md.IndexOf("Job role:");
                    md = md.Substring(b, a - b);
                    md = md.Replace("####", "#").Replace("English", "# About this course");
                    md = StripHTML(md);
                    md = md.Substring(0, md.Length - 2);
                    CreateFile(md, filename);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(b);
                    Console.WriteLine(a);

                    Console.WriteLine(filename);

                }
            }
            excelFileGenerator.dosmth();

            MessageBox.Show("All Finished");



        }

        private void CreateFile(string text, string filename)
        {
            DirectoryInfo di = Directory.CreateDirectory("mdcoursesfiles");

            using (System.IO.StreamWriter file = new System.IO.StreamWriter($"mdcoursesfiles/{filename}"))
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
            filename = filename.Substring(0, filename.IndexOf(" ")).ToLower();
            return filename;
        }

        public static string StripHTML(string input)
        {

            input = Regex.Replace(input, "<.*?>", String.Empty).Replace("Show more", "").Replace("See all instructor-led courses", "");
            //input = Regex.Replace(input, @"\((.*?)spx\)", String.Empty);
            input = Regex.Replace(input, @"\[[^>]*spx\)|\![^>]*svg\)|\[[^>]*browse/\)", String.Empty);

            return Regex.Replace(input, @"# /  +/g\n|# \n", "#");

        }


        public string TakeCourseTitle(string htmlContent)
        {
            var titleStartIndex = htmlContent.IndexOf("<title>") ;
            string title = htmlContent.Substring(titleStartIndex, htmlContent.IndexOf("</title>") - titleStartIndex);

            return title.Substring("<title>".Length + "Course ".Length, title.Length - " -Learn | Microsoft Docs".Length - "<title>".Length - "Course ".Length);
        }

        public void GetAllHrefs(String uri)
        {



            var web = new HtmlWeb();
            web.BrowserTimeout = TimeSpan.FromSeconds(30);
            
            var doc = web.LoadFromBrowser(uri, o =>
            {
               // var webBrowser = (WebBrowser)o;
                

                return o.Contains("card-content-title");
                
            });

            GC.Collect();
            // Show results

            Regex re = new Regex(@">(.*?)<");
            foreach (var title in doc.DocumentNode.SelectNodes(".//div[@class='card-content']"))
            {
                if (title.SelectSingleNode("a[@class='card-content-title']") != null)
                {
                    string url = "https://docs.microsoft.com/en-us" + title.SelectSingleNode("a[@class='card-content-title']").GetAttributeValue("href", String.Empty);
                    hrefslist.Add(url);
                    if (title.SelectSingleNode("ul[@class='tags']") != null)
                    {
                        string result = "";
                        string a = (title.SelectSingleNode("ul[@class='tags']").InnerText);




                        foreach (Match match in re.Matches(a))
                        {
                            if (match.Groups[1].Length != 0)
                            {
                                result = match.Groups[1].Value;
                                break;
                            }
                              /*  if(!match.Groups[1].Equals("Advanced") && !match.Groups[1].Equals("Intermediate") && !match.Groups[1].Equals("Beginner") )
                                    {
                                    result += match.Groups[1] + ", ";
                                } */


                        }
                       // if(result.Length>2) result = result.Substring(0, result.Length - 2);



                        prodDictionary[url] = result;


                    }

                }

            }




        }

        void CreateUrlFile(String filename)
        {
            TextWriter tw = new StreamWriter(filename);
            foreach (String href in hrefslist)
                tw.WriteLine(href);

            tw.Close();
        }

        private void UrlGeneration()
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
            MessageBox.Show("First Step Finished");

        }


    }
}
