using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace myonemap
{
    public partial class Connect
    {

        /*private Command _AddNoteToCurrentSection;

          //**********add note to current section
                _AddNoteToCurrentSection = _MindManager.Commands.Add("myonemap.Connect", "ssst");
                _AddNoteToCurrentSection.ToolTip = "Add a new onenote page to current section";
                _AddNoteToCurrentSection.Caption = "Add new onenote page to current sectin";
                Utils.AddClickHandler(_AddNoteToCurrentSection, AddNoteToCurrentSection);
                Utils.AddUpdateStateHandler(_AddNoteToCurrentSection, AddNoteToCurrentSectionUpdateState);

        public void AddNoteToCurrentSection()
        {
            //see if any onenote windows is open
            var oneApp = new Microsoft.Office.Interop.OneNote.Application();
            string outstr = string.Empty;
            if (oneApp.Windows.Count == 0)
            {
                if (MessageBox.Show("Select a section in OneNote and try again?", "onenot is not open", MessageBoxButtons.YesNo) ==
                    DialogResult.Yes)
                {
                    Process oneProc = new Process();
                    oneProc.StartInfo.FileName = "onenote.exe";
                    //oneProc.StartInfo.Arguments = "/docked";
                    oneProc.Start();
                    Thread.Sleep(300);
                }
            }
            else
            {
                try
                {
                    string pageId = string.Empty;
                    var selTopic = _MindManager.ActiveDocument.Selection.PrimaryTopic;
                    string currentSectionId = string.Empty;
                    currentSectionId = oneApp.Windows[(uint)0].CurrentSectionId;
                    pageId = OnenoteUtils.CreatePage(ref oneApp, currentSectionId, selTopic.Text);
                    //oneApp.CreateNewPage(oneApp.Windows[(uint)0].CurrentSectionId, out pageId, NewPageStyle.npsDefault);

                    string mmGuid = OnenoteUtils.AddOrGetMeta(ref oneApp, pageId,
                        new Meta("mindmanagerguid", Guid.NewGuid().ToString()));
                    string ret = OnenoteUtils.GetLink(ref oneApp, pageId);

                    ret += "&mm-guid={" + mmGuid + "}";
                    //todo: handle situations where the page has not title and the & in the &section-id does not exist
                    Hyperlink hyperlink = selTopic.Hyperlinks.AddHyperlink(ret.Substring(0, ret.IndexOf("section-id")));
                    hyperlink.Arguments = ret.Substring(ret.IndexOf("section-id"));
                    _MindManager.ActiveDocument.Save();
                    oneApp.NavigateTo(pageId);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }
            //create a new page in the currentsection
            //oneApp.CreateNewPage(oneApp.Windows[(uint)0].CurrentSectionId,out pageId, NewPageStyle.npsBlankPageNoTitle);
        }
           private void AddNoteToCurrentSectionUpdateState(ref bool penabled, ref bool pchecked)
        {
            penabled = true;
            /*    var oneApp = new OnenoteApplication();
                if (oneApp.Windows.Count == 0)
                    penabled = false;#1#
        }
         
         
         
         
         
         */



/*
        private void ImportLinks()
        {
            var oneApp = new OnenoteApplication();
            var selTopic = _MindManager.ActiveDocument.Selection.PrimaryTopic;
            if (oneApp.Windows.Count == 0)
            {

                if (
                    MessageBox.Show("Select a section in OneNote and try again?", "onenot is not open",
                        MessageBoxButtons.YesNo) ==
                    DialogResult.Yes)
                {
                    var oneProc = new Process();
                    oneProc.StartInfo.FileName = "onenote.exe";
                    //oneProc.StartInfo.Arguments = "/docked";
                    oneProc.Start();
                    Thread.Sleep(300);
                }
            }
            else
            {
                string currSection = oneApp.Windows[(uint)0].CurrentSectionId;
                var pDialog = new ProgressDialog();
                try
                {

                    if (!string.IsNullOrEmpty(currSection))
                    {
                        var owner = new Win32Window(Process.GetCurrentProcess().MainWindowHandle);

                        pDialog.Show(owner);

                        //pDialog.PlayCursor();
                        pDialog.SetPbarVisiblity(false);

                        //  var links = OnenoteUtils.GetSectionLinks(ref oneApp, currSection);
                        //    pDialog.Message = "writing logs-------";
                        //  links.ForEach(x => Utils.Pt(x.Item1.ToString()));
                        string temp = string.Empty;
                        // links.ForEach(x => temp += "\n" + x.Item1 + "----title: " + x.Item2);
                        //MessageBox.Show(temp);
                        if (selTopic != null)
                        {
                            var pageIds = OnenoteUtils.GetSectionPageIds(ref oneApp, currSection);
                            int count = pageIds.Count();
                            int block = 100 / count;
                            pDialog.Progress = 100 % count;
                            //pDialog.Message = "Importing pages ...";
                            pDialog.Message = Process.GetCurrentProcess().ProcessName.ToLower();
                            pDialog.SetPbarVisiblity(true);
                            foreach (var pageId in pageIds)
                            {
                                var x = OnenoteUtils.CreateLinksTemp(ref oneApp, pageId);
                                try
                                {
                                    pDialog.Progress += block;
                                    Topic topic = selTopic.AddSubTopic(x.Item2);
                                    if (x.Item1.Contains("&section-id"))
                                    {
                                        topic.Hyperlinks.AddHyperlink(x.Item1.Substring(0, x.Item1.IndexOf("&section-id")));
                                        topic.Hyperlink.Arguments = x.Item1.Substring(x.Item1.IndexOf("&section-id"));
                                    }
                                    else
                                    {
                                        topic.Hyperlinks.AddHyperlink(x.Item1.Substring(0, x.Item1.IndexOf("section-id")));
                                        topic.Hyperlink.Arguments = "&" + x.Item1.Substring(x.Item1.IndexOf("section-id"));
                                    }
                                    //_MindManager.ActiveDocument.Save();
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show("Erorr importing sectin" + ex.Message);
                                }
                            }


                        }



                        /*   if (selTopic != null)
                        {
                            int count = links.Count();
                            int block = 100/count;
                            pDialog.Progress = 100%count;
                            pDialog.Message = "Importing pages ...";
                            pDialog.SetPbarVisiblity(true);
                            pDialog.StopCursor();
                            foreach (var x in links)
                            {
                            
                                 try
                                 {
                                     pDialog.Progress += block;
                                    Topic topic = selTopic.AddSubTopic(x.Item2);
                                    if (x.Item1.Contains("&section-id"))
                                    {
                                        topic.Hyperlinks.AddHyperlink(x.Item1.Substring(0, x.Item1.IndexOf("&section-id")));
                                        topic.Hyperlink.Arguments = x.Item1.Substring(x.Item1.IndexOf("&section-id"));
                                    }
                                    else
                                    {
                                        topic.Hyperlinks.AddHyperlink(x.Item1.Substring(0, x.Item1.IndexOf("section-id")));
                                        topic.Hyperlink.Arguments = "&" + x.Item1.Substring(x.Item1.IndexOf("section-id"));
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show("Erorr importing sectin" + ex.Message);
                                }
                                
                            }
                            _MindManager.ActiveDocument.Save();
                        }#1#
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error importing section: " + ex.Message);
                }
                finally
                {
                    pDialog.Close();
                    Utils.ReleaseComObject(oneApp);
                }
            }
        }
        private void ImportLinks2()
        {
            string retstring = string.Empty;
            var onenoteApp = new OnenoteApplication();
            var selectLinkForm = new SelectOneNotePage(ref onenoteApp, HierarchyType.Sections);

            try
            {
                DialogResult dr = selectLinkForm.ShowDialog();
                retstring = SelectOneNotePage.retlink;

                if (!string.IsNullOrEmpty(retstring))
                {
                    var selTopic = _MindManager.ActiveDocument.Selection.PrimaryTopic;
                    var pDialog = new ProgressDialog();
                    try
                    {

                        var owner = new Win32Window(Process.GetCurrentProcess().MainWindowHandle);
                        pDialog.Show(owner);

                        //pDialog.PlayCursor();
                        pDialog.SetPbarVisiblity(false);

                        //  var links = OnenoteUtils.GetSectionLinks(ref oneApp, currSection);
                        //    pDialog.Message = "writing logs-------";
                        //  links.ForEach(x => Utils.Pt(x.Item1.ToString()));
                        string temp = string.Empty;
                        // links.ForEach(x => temp += "\n" + x.Item1 + "----title: " + x.Item2);
                        //MessageBox.Show(temp);
                        if (selTopic != null)
                        {
                            var pageIds = OnenoteUtils.GetSectionPageIds(ref onenoteApp, retstring);
                            int count = pageIds.Count();
                            int block = 100 / count;
                            pDialog.Progress = 100 % count;
                            //pDialog.Message = "Importing pages ...";
                            pDialog.Message = Process.GetCurrentProcess().ProcessName.ToLower();
                            pDialog.SetPbarVisiblity(true);
                            foreach (var pageId in pageIds)
                            {
                                var x = OnenoteUtils.CreateLinksTemp(ref onenoteApp, pageId);
                                try
                                {
                                    pDialog.Progress += block;
                                    Topic topic = selTopic.AddSubTopic(x.Item2);
                                    if (x.Item1.Contains("&section-id"))
                                    {
                                        topic.Hyperlinks.AddHyperlink(x.Item1.Substring(0, x.Item1.IndexOf("&section-id")));
                                        topic.Hyperlink.Arguments = x.Item1.Substring(x.Item1.IndexOf("&section-id"));
                                    }
                                    else
                                    {
                                        topic.Hyperlinks.AddHyperlink(x.Item1.Substring(0, x.Item1.IndexOf("section-id")));
                                        topic.Hyperlink.Arguments = "&" + x.Item1.Substring(x.Item1.IndexOf("section-id"));
                                    }
                                    //_MindManager.ActiveDocument.Save();
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show("Erorr importing sectin" + ex.Message);
                                }
                            }


                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error importing section: " + ex.Message);
                    }
                    finally
                    {
                        pDialog.Close();
                    }
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                //selectLinkForm.Dispose();
                Utils.ReleaseComObject(onenoteApp);
                retstring = string.Empty;
            }


        }*/
    }
}
