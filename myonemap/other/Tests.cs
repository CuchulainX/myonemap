using System;
using System.Collections;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using Mindjet.MindManager.Interop;
using myonemap.onenote;
using OnenoteApplication = Microsoft.Office.Interop.OneNote.Application;
using MindmanagerApplication = Mindjet.MindManager.Interop.Application;
using Control = Mindjet.MindManager.Interop.Control;
using Utilities = Mindjet.MindManager.Interop.Utilities;
namespace myonemap
{
    public partial class Connect
    {
/*
        private string section1_id;
        private string section2_id;
        private string notebookId;
        public void RunTests()
        {
            //make sure testmap is loaded
            if (TestMapLoadedAndActive())
            {
                //  InitializeForTests();
/*

                ImportLinks();
                Thread.Sleep(2000);
                // change section name
                ChangeSectionName();
                Thread.Sleep(2000);
                // sync hyperlinks with onemap
                SyncDocumentWithOnenote();
#1#

                //move page
                MoveSectionPages();
            }
        }

        private void MoveSectionPages()
        {
            var app = new OnenoteApplication();
            try
            {
                string outstr = string.Empty;
                app.GetHierarchy(notebookId, HierarchyScope.hsSections, out outstr);
                var doc = XDocument.Parse(outstr);
                var ns = doc.Root.Name.Namespace;
                var section = doc.Descendants(ns + "Section").FirstOrDefault(x => x.Attribute("name").Value == "s1");
                var targetSection =
                    doc.Descendants(ns + "Section").FirstOrDefault(x => x.Attribute("name").Value == "section2");
                if (section != null)
                {
                    var q = from node in section.Descendants(ns + "Page")
                        select node;
                    foreach (var page in q)
                    {
                        page.Parent.SetValue(targetSection);
                    }
                }
                app.UpdateHierarchy(doc.ToString());
            }
            catch (COMException exception)
            {
                var message = OneNoteHresultDescriptions.GetErrorDescription(exception.ErrorCode);
            }
            catch (Exception exception)
            {

            }
            finally
            {
                Utils.ReleaseComObject(app);
            }
        }

        private void ChangeSectionName()
        {
            var app = new OnenoteApplication();
            try
            {
                string outstr = string.Empty;
                app.GetHierarchy(notebookId, HierarchyScope.hsSections, out outstr);
                var doc = XDocument.Parse(outstr);
                var ns = doc.Root.Name.Namespace;
                var section = doc.Descendants(ns + "Section").FirstOrDefault(x => x.Attribute("name").Value == "section1");
                if (section != null)
                {
                    section.Attribute("name").SetValue("s1");
                }
                app.UpdateHierarchy(doc.ToString());
            }
            catch (COMException exception)
            {
                var message = OneNoteHresultDescriptions.GetErrorDescription(exception.ErrorCode);
            }
            catch (Exception exception)
            {

            }
            finally
            {
                Utils.ReleaseComObject(app);
            }
        }

        private void InitializeForTests()
        {
            try
            {
                //delete all topics in the testmap
                DeleteAllTopics();

                //add test topics
                AddTestTopics();

                //add links to topics
                //AddLinkToTopis();

                //open a onenote test notebook
                //check for notebook by name
                if (NotbookCurrentlyOpen("testnotebook"))
                {
                    string path = @"D:\temp\myonemap\testnotebook\testnotebook\";
                    notebookId = TryGetNotebookId("testnotebook");
                    DeleteAllSections(notebookId);

                    //create some test sections
                    section1_id = createSection(path, "section1");
                    section2_id = createSection(path, "section2");
                    var section3_id = createSection(path, "section3");
                    var sectionr_id = createSection(path, "section4");
                    var page1_id = createTestPage(section1_id, "page1");
                    var page2_id = createTestPage(section1_id, "page2");
                    for (int i = 0; i < 10; i++)
                    {
                        createTestPage(section1_id, i.ToString());
                    }
                    // import some links
                }
            }
            catch (Exception exception)
            {

            }
        }

        private object createTestPage(string sectionId, string pageName)
        {
            var app = new OnenoteApplication();
            string retId = string.Empty;
            try
            {
                app.CreateNewPage(sectionId, out retId);
            }
            catch (COMException exception)
            {
                var message = OneNoteHresultDescriptions.GetErrorDescription(exception.ErrorCode);
            }
            catch (Exception exception)
            {

            }
            finally
            {
                Utils.ReleaseComObject(app);
            }
            return retId;
        }

        private string createSection(string path, string sectionName)
        {
            var app = new OnenoteApplication();
            string retId = string.Empty;
            try
            {
                app.OpenHierarchy(path + sectionName + ".one", string.Empty, out retId, CreateFileType.cftSection);
            }
            catch (COMException exception)
            {
                var message = OneNoteHresultDescriptions.GetErrorDescription(exception.ErrorCode);
            }
            catch (Exception exception)
            {

            }
            finally
            {
                Utils.ReleaseComObject(app);
            }
            return retId;
        }

        private void DeleteAllSections(string notebookId)
        {
            var app = new OnenoteApplication();
            try
            {
                string xmlstr;
                app.GetHierarchy(notebookId, HierarchyScope.hsSections, out xmlstr);
                var doc = XDocument.Parse(xmlstr);
                var ns = doc.Root.Name.Namespace;
                var q = from node in doc.Descendants(ns + "Section")
                        select node;
                foreach (var sectionId in q)
                {
                    app.DeleteHierarchy(sectionId.Attribute("ID").Value);
                }
            }
            catch (COMException exception)
            {
                var message = OneNoteHresultDescriptions.GetErrorDescription(exception.ErrorCode);
            }
            catch (Exception exception)
            {

            }
            finally
            {
                Utils.ReleaseComObject(app);
            }
        }
        private void DeleteSection(string sectionId)
        {
            var app = new OnenoteApplication();
            try
            {
                string xmlstr;
                app.DeleteHierarchy(sectionId);
            }
            catch (COMException exception)
            {
                var message = OneNoteHresultDescriptions.GetErrorDescription(exception.ErrorCode);
            }
            catch (Exception exception)
            {

            }
            finally
            {
                Utils.ReleaseComObject(app);
            }
        }
        private void DeleteAllPages(string path, string sectionId)
        {
            var app = new OnenoteApplication();
            try
            {
                string xmlstr;
                // app.OpenHierarchy(path,sectionId, out xmlstr);
                app.OpenHierarchy(path, string.Empty, out xmlstr);

                var doc = XDocument.Parse(xmlstr);
                var ns = doc.Root.Name.Namespace;
                var q = from node in doc.Descendants(ns + "Section").Descendants(ns + "Page") select node.Attribute("ID").Value;
                foreach (var page in q)
                {
                    app.DeleteHierarchy(page);
                }
            }
            catch (COMException exception)
            {
                var message = OneNoteHresultDescriptions.GetErrorDescription(exception.ErrorCode);
            }
            catch (Exception exception)
            {

            }
            finally
            {

            }
        }

        /* public static void DeleteAllSections(string notebookid)
         {
             var app = new OnenoteApplication();
             try
             {
                 string xmlStr = String.Empty;
                 app.GetHierarchy(notebookid, HierarchyScope.hsSections, out xmlStr);
                 var doc = XDocument.Parse(xmlStr);
                 var ns = doc.Root.Name.Namespace;
                 var sectionIdList = new List<string>();
                 foreach (var sectionNode in from node in doc.Descendants(ns + "Section") select node)
                 {
                    sectionIdList.Add(sectionNode.Attribute("ID").Value);
                 }
                 foreach (var sectionId in sectionIdList)
                 {
                     app.DeleteHierarchy(sectionId);
                 }
             }
             catch (COMException exception)
             {
                 var desc = OneNoteHresultDescriptions.GetErrorDescription(exception.ErrorCode);
             }
             catch (Exception exception)
             {

             }
             finally
             {

             }
         }#1#

        private string TryGetsectionId(string notebookName, string sectionName)
        {
            var app = new OnenoteApplication();
            string xmlStr = string.Empty;
            string retVal = string.Empty;
            try
            {
                var notebookId = TryGetNotebookId(notebookName);
                if (!string.IsNullOrEmpty(notebookId))
                {
                    app.GetHierarchy(notebookId, HierarchyScope.hsSections, out xmlStr);
                    var doc = XDocument.Parse(xmlStr);
                    var ns = doc.Root.Name.Namespace;
                    var q = from node in doc.Descendants(ns + "Section")
                            where node.Attribute("name").Value == sectionName
                            select node.Attribute("ID").Value;
                    if (q.Any())
                        retVal = q.First();
                }
            }
            catch (COMException exception)
            {
                var message = OneNoteHresultDescriptions.GetErrorDescription(exception.ErrorCode);
            }
            catch (Exception exception)
            {

            }
            finally
            {
                Utils.ReleaseComObject(app);
            }

            return retVal;
        }

        private string TryGetNotebookId(string notebookName)
        {
            var app = new OnenoteApplication();
            string ret = string.Empty;
            try
            {
                string notebookXml;
                app.GetHierarchy(null, HierarchyScope.hsNotebooks, out notebookXml);
                var doc = XDocument.Parse(notebookXml);
                var ns = doc.Root.Name.Namespace;
                foreach (var notebook in from node in doc.Descendants(ns + "Notebook") select node)
                {
                    if (notebook.Attribute("name").Value == notebookName)
                    {
                        ret = notebook.Attribute("ID").Value;
                        break;
                    }
                }
            }
            catch (Exception e)
            {

            }
            finally
            {
                Utils.ReleaseComObject(app);
            }
            return ret;
        }

        private string TryOpenSectionByPath(string path)
        {
            var app = new OnenoteApplication();
            string notebookId;
            try
            {
                app.OpenHierarchy(path, string.Empty, out notebookId, CreateFileType.cftNone);
                return notebookId;
            }
            catch (COMException exception)
            {
                var errDescrption = OneNoteHresultDescriptions.GetErrorDescription(exception.ErrorCode);
            }
            catch (Exception exception)
            {

            }
            finally
            {
                Utils.ReleaseComObject(app);
            }
            return string.Empty;
        }

        private bool NotbookCurrentlyOpen(string testnotebook)
        {
            var app = new OnenoteApplication();
            bool ret = false;
            try
            {
                string notebookXml;
                app.GetHierarchy(null, HierarchyScope.hsNotebooks, out notebookXml);
                var doc = XDocument.Parse(notebookXml);
                var ns = doc.Root.Name.Namespace;
                foreach (var notebook in from node in doc.Descendants(ns + "Notebook") select node)
                {
                    if (notebook.Attribute("name").Value == testnotebook)
                    {
                        ret = true;
                        break;
                    }
                }
            }
            catch (Exception e)
            {

            }
            finally
            {
                Utils.ReleaseComObject(app);
            }
            return ret;
        }

        private void AddLinkToTopis()
        {
            try
            {
                IEnumerator enumerator = _MindManager.ActiveDocument.Range(MmRange.mmRangeAllTopics, false).GetEnumerator();
                while (enumerator.MoveNext())
                {
                    if (enumerator.Current as Topic != null)
                    {
                        if (enumerator.Current as Topic != _MindManager.ActiveDocument.CentralTopic)
                        {
                            (enumerator.Current as Topic).Hyperlinks.AddHyperlink("http://www.google.com");
                            (enumerator.Current as Topic).Hyperlinks.AddHyperlink("http://www.yahoo.com");
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void AddTestTopics()
        {
            try
            {
                _MindManager.ActiveDocument.CentralTopic.AddSubTopic("test subtopic1");
                _MindManager.ActiveDocument.CentralTopic.AddSubTopic("2");
                _MindManager.ActiveDocument.CentralTopic.AddSubTopic("3");
                _MindManager.ActiveDocument.CentralTopic.AddSubTopic("4");
                _MindManager.ActiveDocument.CentralTopic.AddSubTopic("5");
            }
            catch (Exception ex)
            {

            }

        }

        public void DeleteAllTopics()
        {
            try
            {
                IEnumerator enumerator = _MindManager.ActiveDocument.Range(MmRange.mmRangeAllTopics, false).GetEnumerator();
                while (enumerator.MoveNext())
                {
                    if (enumerator.Current as Topic != null)
                    {
                        if (enumerator.Current as Topic != _MindManager.ActiveDocument.CentralTopic)
                        {
                            (enumerator.Current as Topic).Delete();
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        public bool TestMapLoadedAndActive()
        {
            var ret = true;
            if (_MindManager.ActiveDocument.Name != "testmap.mmap")
                ret = false;
            return ret;
        }*/
    }
}