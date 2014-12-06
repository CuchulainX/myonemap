using System.Windows.Forms;

namespace myonemap
{
    public class Link
    {
        public string BasePath { get; private set; }
        public string Title { get; private set; }
        public string SectionId { get; private set; }
        public string PageId { get; private set; }
        public string Mmguid { get; private set; }

        public Link(string basePath = "", string title = "", string sectionId = "", string pageId = "",
            string mmguid = "")
        {
            BasePath = basePath;
            Title = title;
            SectionId = sectionId;
            PageId = pageId;
            Mmguid = mmguid;
        }
    }

    public static class LinkExtensions
    {
        public static Link AsLink(this string source)
        {
            string BasePath = GetBasePath(source);
            string Title = GetTitle(source);
            //SectionId = GetSeciontId(input);
            //PageId = GetPageId(input);
            //Mmguid = GetMmguid(input);
            return new Link(basePath: BasePath,title:Title);
        }

        private  static string GetMmguid(string input)
        {
            string ret = string.Empty;
            return ret;
        }

        private static string GetPageId(string input)
        {
            string ret = string.Empty;
            if (input.Contains("page-id"))
            {
                input = input.Substring(input.IndexOf("page-id"));
                input = input.Substring(0, input.IndexOf('&'));
                input.Replace("{", "").Replace("}", "").Replace("&","").Replace("page-id=", "");
                ret = input;
            }
            return ret;
        }

        private static string GetSectionId(string input)
        {
            string ret = string.Empty;
            if (input.Contains("section-id"))
            {
                input = input.Substring(input.IndexOf("section-id"));
                input = input.Substring(0, input.IndexOf('&'));
                input.Replace("{", "").Replace("}", "").Replace("&","").Replace("section-id=", "");
                ret = input;
            }
            return ret;
        }

        private static string GetTitle(string input)
        {
            string ret = string.Empty;
            if (input.Contains("onenote:///"))
            {
                input = input.Replace("onenote:///", "");
                if (input.Contains("#") && input.Contains("&section-id"))
                {
                    var temp = input.Substring(input.IndexOf("&section-id"));
                    input.Replace(temp, "");
                    ret = input.Substring(input.IndexOf("#"));
                }
            }
            return ret;
        }

        private static string GetBasePath(string input)
        {
            string ret = string.Empty;
            if (input.Contains("onenote:///"))
            {
                input = input.Replace("onenote:///", "");
                if (input.Contains("#&section-id"))
                {
                    ret = input.Substring(0, input.IndexOf("#&section-id"));
                }
                else if (input.Contains("#section-id"))
                {
                    ret = input.Substring(0, input.IndexOf("#section-id"));
                }
                else if (input.Contains("#") && input.Contains("section-id"))
                {
                    if (input.IndexOf("#") < input.IndexOf("section-id"))
                    {
                        ret = input.Substring(0, input.IndexOf("#"));
                    }
                }
            }
            return ret;
        }


    }
}