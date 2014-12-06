using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;

namespace myonemap.core.onenote
{
    public static class OnenoteUtils
    {
        /// <summary>
        /// check onenote page to see if it has any meta attributes
        /// </summary>
        /// <param name="pageId"></param>
        /// <returns></returns>
        public static bool HasMetaAttribute(string pageId)
        {
            string notebookXml;
            //onenoteApp.GetPageContent(pageId, out notebookXml);
            OneNote.Instance().GetPageContent(pageId, out notebookXml);
            var doc = XDocument.Parse(notebookXml);
            var ns = doc.Root.Name.Namespace;
            bool retVal = doc.Elements(ns + "Page").Elements(ns + "Meta").FirstOrDefault() == null;
            return !retVal;
        }

        /// <summary>
        /// checks onenote page to see if it has a meta attribute with a specific name
        /// </summary>
        /// <param name="pageId"></param>
        /// <param name="metaName"></param>
        /// <returns></returns>
        public static bool HasMetaAttribWithName(string pageId, string metaName)
        {

            if (HasMetaAttribute(pageId))
            {
                string notebookXml;
                //onenoteApp.GetPageContent(pageId, out notebookXml);
                OneNote.Instance().GetPageContent(pageId, out notebookXml);
                var doc = XDocument.Parse(notebookXml);
                var ns = doc.Root.Name.Namespace;
                bool retVal =
                    doc.Elements(ns + "Page")
                        .Elements(ns + "Meta")
                        .FirstOrDefault(x => x.Attribute("name").Value == metaName) == null;
                return !retVal;
            }
            return false;
        }

        /// <summary>
        /// Get the value of a meta attribute with specific name if it exists
        /// </summary>
        /// <param name="pageId"></param>
        /// <param name="metaName"></param>
        /// <returns></returns>
        public static Meta GetMeta(string pageId, string metaName)
        {
            if (HasMetaAttribWithName(pageId, metaName))
            {
                string notebookXml;
                //onenoteApp.GetPageContent(pageId, out notebookXml);
                OneNote.Instance().GetPageContent(pageId, out notebookXml);
                var doc = XDocument.Parse(notebookXml);
                var ns = doc.Root.Name.Namespace;
                var q =
                    doc.Elements(ns + "Page")
                        .Elements(ns + "Meta")
                        .Where(x => x.Attribute("name").Value == metaName)
                        .Select(x => new Meta(x.Attribute("name").Value, x.Attribute("content").Value)).Single();
                return q;
            }
            return null;
        }

        /// <summary>
        /// Get a meta attribute based on info in the Meta object and creates it if it doesn't exist
        /// </summary>
        /// <param name="pageId"></param>
        /// <param name="meta"></param>
        /// <returns></returns>
        public static string AddOrGetMeta(string pageId, Meta meta)
        {
            if (HasMetaAttribWithName(pageId, meta.Name))
            {
                var retMeta = GetMeta(pageId, meta.Name);
                return retMeta.Content;
            }
            try
            {
                string notebookXml;
                //onenoteApp.GetPageContent(pageId, out notebookXml);
                OneNote.Instance().GetPageContent(pageId, out notebookXml);
                var doc = XDocument.Parse(notebookXml);
                var ns = doc.Root.Name.Namespace;
                var existingPageId = doc.Root.Attribute("ID").Value;

                var page = new XDocument(new XElement(ns + "Page",
                    new XElement(ns + "Meta",
                        new XAttribute("name", meta.Name.ToString()),
                        new XAttribute("content", meta.Content.ToString())
                        )));
                page.Root.SetAttributeValue("ID", existingPageId);
                doc.Root.Document.Element(ns + "Page").AddFirst(page.ToString());
                //doc.Root.Element("Page").AddFirst(page.ToString());
                OneNote.Instance().UpdatePageContent(page.ToString(), DateTime.MinValue);
                return meta.Content;
            }
            catch (COMException ceException)
            {
                //Console.WriteLine(ceException.Message.ToString());
                //throw new Exception("error adding meta tag", ceException);
                throw new OneMapException("error adding meta tag", ceException);
            }
            catch (Exception ex)
            {
                throw new OneMapException("error in AddOrGetMeta", ex);
            }

        }

        /// <summary>
        /// get a list of pageIds in a section
        /// </summary>
        /// <param name="currSection"></param>
        /// <returns></returns>
        public static List<string> GetSectionPageIds(string currSection)
        {
            string outstr = string.Empty;
            List<string> ret = new List<string>();
            OneNote.Instance().GetHierarchy(currSection, HierarchyScope.hsPages, out outstr);
            try
            {
                if (!string.IsNullOrEmpty(outstr))
                {
                    var doc = XDocument.Parse(outstr);
                    var ns = doc.Root.Name.Namespace;
                    doc.Descendants(ns + "Page").ToList().ForEach(x => ret.Add(x.Attribute("ID").Value));
                    return ret;
                }
            }
            catch (Exception ex)
            {
                throw new OneMapException("GetSectionPageIds-> error getting section", ex);
            }
            return null;
        }

        /// <summary>
        /// create a a link for mindmanager hyperlink
        /// </summary>
        /// <param name="selectedItem"></param>
        /// <returns></returns>
        public static string GetMindManagerLink(string selectedItem)
        {
            string mmGuid = OnenoteUtils.AddOrGetMeta(selectedItem,
                new Meta("mindmanagerguid", Guid.NewGuid().ToString()));
            string retstring = OnenoteUtils.GetLink(selectedItem);
            retstring += "&mm-guid={" + mmGuid + "}";
            return retstring;
        }
        /// <summary>
        /// find the page that contains the mindmanager guid
        /// </summary>
        /// <param name="mmguid"></param>
        /// <returns></returns>
        public static string GetHyperLinkBymmGuid(string mmguid)
        {
            //OnenoteApplication onApplication = new OnenoteApplication();
            string outstr = string.Empty;
            string retString = string.Empty;
            try
            {
                OneNote.Instance().FindMeta(string.Empty, "mindmanagerguid", out outstr);
                var doc = XDocument.Parse(outstr);
                var ns = doc.Root.Name.Namespace;

                var pageNode =
                    from page in doc.Descendants(ns + "Notebook").Descendants(ns + "Section").Descendants(ns + "Page")
                    let isInRecycleBin = GetAttributeValue(page, "isInRecycleBin", string.Empty) == string.Empty
                    from meta in page.Descendants(ns + "Meta")
                    where isInRecycleBin && meta.Attribute("content").Value == mmguid
                    select page.Attribute("ID").ToString();

                //todo: hanadle reults with zero results
                if (pageNode.Count() != 0)
                {
                    var pageNodeId = pageNode.Single().ToString();

                    pageNodeId = pageNodeId.Replace("ID=\"", "").Replace("\"", "");
                    retString = GetHyperLinkByObjectId(pageNodeId);
                }
            }
            catch (COMException ex)
            {
                throw new OneMapException(OneNoteHresultDescriptions.GetErrorDescription(ex.ErrorCode),
                    ex, ex.GetBaseException());
            }
          
            return retString;
        }

        public static bool WSearchIsOn()
        {
            var service = ServiceController.GetServices().FirstOrDefault(x => x.ServiceName == "WSearch");
            var ret = false;
            if (service != null)
                ret = service.Status == ServiceControllerStatus.Running;
            return ret;
        }

        public static string GetAttributeValue(XElement el, string attributeName, string defaultValue)
        {
            if (el.Attribute(attributeName) != null)
            {
                return (string) el.Attribute(attributeName);
            }
            return defaultValue;
        }

        public static string GetHyperLinkByObjectId(string pageId)
        {
            string hyperlinkAddress = string.Empty;
            //      OnenoteApplication onApplication = new OnenoteApplication();

            OneNote.Instance().GetHyperlinkToObject(pageId, null, out hyperlinkAddress);
            //OneNote.Instance().GetHyperlinkToObject(null,pageId, out hyperlinkAddress);

            return hyperlinkAddress;
        }

        public static string GetLink(string pageId)
        {
            string ret = string.Empty;
            OneNote.Instance().GetHyperlinkToObject(pageId, null, out ret);
            return ret;
        }



        //************ samples and functions to be used later ***************
        /*
        #region enumerations
        public static int indentation = 0;

         *   public static bool PageExist(string address)
        {
            string xmlout = string.Empty;
            string pageAddress = string.Empty;
            bool retvalue = false;
            try
            {
                //address = address.Substring(0, address.IndexOf("&mm-guid"));
                
                string pageid = ExtractPageGuidFromAddress(address);
                
                if (!string.IsNullOrEmpty(pageid))
                    
                {
                    OneNote.Instance().GetPageContent(pageid,out xmlout);
                    //pageAddress =OnenoteUtils.EnumerateAllPages().FirstOrDefault(x => x.Contains(pageid));
                }
                    //pageAddress = FindPageById(pageid);
                if (!string.IsNullOrEmpty(pageAddress))
                    retvalue = true;

            }
            catch (COMException ex)
            {
                 var msg = Marshal.GetExceptionForHR((int)ex.ErrorCode).Message;
                var error = OneNoteHresultDescriptions.GetErrorDescription(ex.ErrorCode);
            }
            catch (Exception e)
            {
                throw new OneMapException("Error inside PageExist ",e);
            }
            //extract the section id

            //use enumerate section for page
            return retvalue;
        }
   
        public static IEnumerable<string> EnumerateGroupSectionsById(string id)
        {
            indentation += 3;
            //var onenoteApp = new OnenoteApplication();
            string notebookXml;
            OneNote.Instance().GetHierarchy(id, HierarchyScope.hsChildren, out notebookXml);

            var doc = XDocument.Parse(notebookXml);
            var ns = doc.Root.Name.Namespace;
            foreach (var sectionGroup in from node in doc.Descendants(ns + "SectionGroup") select node)
            {
                var sectionGroupId = sectionGroup.Attribute("ID").Value;
                if (sectionGroupId != id)
                {
                    Console.WriteLine(IndentWrite() + sectionGroup.Attribute("name").Value);
                    EnumerateGroupSectionsById(sectionGroupId);
                }
            }
            foreach (var section in from node in doc.Descendants(ns + "Section") select node)
            {
                var sectionId = section.Attribute("ID").Value;
                Console.WriteLine(IndentWrite() + section.Attribute("name").Value);
                foreach (var page in EnumerateGroupById(sectionId))
                {
                    yield return page;
                }
            }

            indentation -= 3;
        }

        public static IEnumerable<string> EnumerateGroupById(string id)
        {
            indentation += 3;
          //  var onenoteApp = new OnenoteApplication();
            string notebookXml;
            OneNote.Instance().GetHierarchy(id, HierarchyScope.hsChildren, out notebookXml);
            var doc = XDocument.Parse(notebookXml);
            var ns = doc.Root.Name.Namespace;
            foreach (var section in from node in doc.Descendants(ns + "Section") select node)
            {
                section.GetHashCode();
                var sectionId = section.Attribute("ID").Value;
                if (sectionId != id)
                {
                    Console.WriteLine(IndentWrite() + section.Attribute("name").Value);
                    EnumerateGroupById(sectionId);
                }
            }
            foreach (var page in from node in doc.Descendants(ns + "Page") select node)
            {
                Console.WriteLine(IndentWrite() + "-----" + page.Attribute("name").Value);
                yield return page.Attribute("ID").Value;
            }
            indentation -= 3;
        }

        public static IEnumerable<string> EnumerateAllPages()
        {
         //   var onenoteApp = new OnenoteApplication();
            string notebookXml;
            OneNote.Instance().GetHierarchy(null, HierarchyScope.hsNotebooks, out notebookXml);
            var doc = XDocument.Parse(notebookXml);
            var ns = doc.Root.Name.Namespace;
            foreach (var notebook in from node in doc.Descendants(ns + "Notebook") select node)
            {
                foreach (var page in EnumerateNoteBookById(notebook.Attribute("ID").Value))
                {
                    yield return page;
                }
            }
        }

        public static IEnumerable<string> EnumerateNoteBookById(string id)
        {
            indentation += 3;
          //  var onenoteApp = new OnenoteApplication();
            string notebookXml;
            OneNote.Instance().GetHierarchy(id, HierarchyScope.hsSections, out notebookXml);
            var doc = XDocument.Parse(notebookXml);
            var ns = doc.Root.Name.Namespace;
            foreach (var sectionGroup in from node in doc.Descendants(ns + "SectionGroup") select node)
            {
                Console.WriteLine(IndentWrite() + sectionGroup.Attribute("name").Value);
                var ret = EnumerateGroupSectionsById(sectionGroup.Attribute("ID").Value);
                foreach (var item in ret)
                {
                    yield return item;
                }
            }
            foreach (var section in from node in doc.Descendants(ns + "Section") select node)
            {
                Console.WriteLine(IndentWrite() + section.Attribute("name").Value);
                var ret = EnumerateGroupById(section.Attribute("ID").Value);
                foreach (var item in ret)
                {
                    yield return item;
                }
            }
            indentation -= 3;
        }
        #endregion

        public static string IndentWrite()
        {
            var sb = new StringBuilder();
            for (int i = 0; i < indentation; i++)
            {
                sb.Append(" ");
            }
            return sb.ToString();
        }

        public static string FindPageById(string pageId)
        {
            string ret = String.Empty;
            var col = EnumerateAllPages();
            foreach (var page in col)
            {
                if (ExtractPageGuid(page) == pageId)
                {
                    ret = page;
                }
            }
            return ret;
        }

        public static string ExtractPageGuid(string pageGuid)
        {
            var ret = pageGuid.Split(new[] {'{', '}'});
            if (ret.Length > 0)
                return ret[1];
            return String.Empty;
        }

        public static string ExtractPageGuidFromAddress(string address)
        {
            if (address.Contains("&end"))
                address = address.Substring(0, address.IndexOf("&end"));
            if (address.Contains("&page-id"))
            {
                address = address.Substring(address.IndexOf("&page-id")).Replace("&page-id=", "");
                address = address.Replace('{', ' ').Replace('}', ' ');
            }
            Guid success = Guid.Empty;
            if (Guid.TryParse(address, out success))
                return address;
            else
                return string.Empty;
        }
        static string GetEntireHierarchy()
        {
            string strXML = string.Empty;
           // OnenoteApplication onApplication = new OnenoteApplication();

            OneNote.Instance().GetHierarchy(null,
                Microsoft.Office.Interop.OneNote.HierarchyScope.hsPages, out strXML);
            //Clipboard.SetText(strXML);
            return strXML;
        }
     


        public static string CreatePage(string sectionId, string pageName)
        {
            //todo: handle exceptions
            // Create the new page
            string pageId;
            OneNote.Instance().CreateNewPage(sectionId, out pageId, NewPageStyle.npsBlankPageWithTitle);

            // Get the title and set it to our page name
            string xml;
            OneNote.Instance().GetPageContent(pageId, out xml, PageInfo.piAll);
            var doc = XDocument.Parse(xml);
            var ns = doc.Root.Name.Namespace;
            //Title is inside a <one:T> ...pagetitle....</one:T>
            //First() because it is first item on the page.
            var title = doc.Descendants(ns + "T").First();
            title.Value = pageName;

            // Update the page
            OneNote.Instance().UpdatePageContent(doc.ToString());

            return pageId;
        }
        

        public static List<Tuple<string,string>> GetSectionLinks( string sectionId)
        {
            List<Tuple<string,string>> links = new List<Tuple<string,string>>();
            var sectionPageIds = GetSectionPageIds( sectionId);
            try
            {
                foreach (var pageId in sectionPageIds)
                {
                    //var mmGuid = AddOrGetMeta(ref oneApp,pageId,new Meta("mindmanagerguid",Guid.NewGuid().ToString()));
                    //string ret = OnenoteUtils.GetLink(ref oneApp,pageId);
                    //string ret = OnenoteUtils.CreateLinkForMindjet(ref oneApp, pageId);
                    //string pageTitle = string.Empty;
                    //if (ret.Contains(".one#") && ret.Contains("&section"))
                    //{
                    //    int length = ret.IndexOf("&section") - ret.IndexOf("#");
                    //    pageTitle = ret.Substring(ret.IndexOf("#"), length).Replace("#", "");
                    //}
                    //ret += "&mm-guid={" + mmGuid + "}";
                    //Tuple<string,string> tup= new Tuple<string, string>(ret,pageTitle);
                    //links.Add(tup);
             
                    links.Add(CreateLinksTemp(pageId));
                }
                return links;
            }
            catch (Exception ex)
            {
                throw new OneMapException("GetSectionLinks-> Error getting section links",ex);
            }
            
        }


        public static Tuple<string,string> CreateLinksTemp(
            string pageId)
        {
            
            try
            {
                string mmGuid = OnenoteUtils.AddOrGetMeta( pageId,
                    new Meta("mindmanagerguid", Guid.NewGuid().ToString()));
                string ret = OnenoteUtils.GetLink(pageId);
                string pageTitle = string.Empty;


                if (ret.Contains(".one#") && ret.Contains("&section"))
                {
                    int length = ret.IndexOf("&section") - ret.IndexOf("#");
                    pageTitle = ret.Substring(ret.IndexOf("#"), length).Replace("#", "");
                }
                ret += "&mm-guid={" + mmGuid + "}";
                return new Tuple<string, string>(ret,pageTitle);
            }
            catch (Exception ex)
            {
                throw new OneMapException("CreateLinksTemp",ex );
            }
            
        }

        public static string CreateLinkForMindjet(
            string pageId)
        {
            
            try
            {
                string mmGuid = OnenoteUtils.AddOrGetMeta( pageId,
                    new Meta("mindmanagerguid", Guid.NewGuid().ToString()));
                string ret = OnenoteUtils.GetLink( pageId);
                string pageTitle = string.Empty;
                if (ret.Contains(".one#") && ret.Contains("&section"))
                {
                    int length = ret.IndexOf("&section") - ret.IndexOf("#");
                    pageTitle = ret.Substring(ret.IndexOf("#"), length).Replace("#", "");
                }
                ret += "&mm-guid={" + mmGuid + "}";
                return ret;
            }
            catch (Exception ex)
            {
                throw new OneMapException("CreateLinkForMindjet()->error creating mindjet links",ex);
            }
            
        }

        private static InvalidLinkError _ValidateOnenoteLink(string onenoteLink)
        {
            //format of onenote link
            // onenote:///basestrpath#title&section-id={}&page-id={}
            if (!onenoteLink.StartsWith("onenote:///"))
            {
                //return false;
                return InvalidLinkError.onenotepart;
            }
            onenoteLink = onenoteLink.Substring(0, "onenote:///".Length);
            if (!onenoteLink.Contains(".one"))
            {
                //return false;
                return InvalidLinkError.doesntcotainoneextension;
            }
            else
            {
                var basePath = onenoteLink.Substring(0, onenoteLink.IndexOf(".one",StringComparison.Ordinal) + ".one".Length);
                if (!Uri.IsWellFormedUriString(basePath, UriKind.Absolute))
                    return InvalidLinkError.basepath;
                onenoteLink.Replace(basePath, "");
                if (!onenoteLink.Contains("section-id"))
                    return InvalidLinkError.NotSectionId;
                if (!onenoteLink.Contains("page-id"))
                    return InvalidLinkError.NoPageId;
            }
            return InvalidLinkError.success;
        }

        private static InvalidLinkError _ValidateOnenoteLinkWithExtraInfo(string link)
        {
            if (_ValidateOnenoteLink(link) == InvalidLinkError.success)
            {
                if (link.Contains("mm-guid"))
                    return InvalidLinkError.success;
            }
            return InvalidLinkError.NommGuid;
        }

        private static string MapEnum(InvalidLinkError error)
        {
            return Enum.GetName(typeof (InvalidLinkError), error);
        }

        public static string ValidateOnenoteLink(string link)
        {
            return MapEnum(_ValidateOnenoteLink(link));
        }

        public static string ValidateOnenoteLinkWithExtraInfo(string link)
        {
            return MapEnum(_ValidateOnenoteLinkWithExtraInfo(link));
        }

      
    }

    
    public enum InvalidLinkError
    {
        onenotepart,
        doesntcotainoneextension,
        basepath,
        NotSectionId,
        NoPageId,
        success,
        NommGuid
    }
    */
    }
}