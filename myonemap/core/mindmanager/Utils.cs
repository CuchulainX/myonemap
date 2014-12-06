using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Mindjet.MindManager.Interop;
using myonemap.core.onenote;
using myonemap.forms;
using myonemap.structs;
using Utilities;
using Application = Mindjet.MindManager.Interop.Application;

namespace myonemap.core.mindmanager
{
    public static class Utils
    {
        public static ribbonTab GetRibbonTabByName(ref Application app, string ribbonTabName)
        {
            Contract.Requires(string.IsNullOrEmpty(ribbonTabName)==false);
            var retRibbon = app.Ribbon.Tabs.AsEnumerable().ToList().FirstOrDefault(x => x.DisplayName == ribbonTabName);
            if (retRibbon != null)
            {
                return (retRibbon as ribbonTab);
            }
            return null;
        }

        public static void Pt(string s_Message)
        {
            /*
            System.Diagnostics.Trace.WriteLine("============================");
            System.Diagnostics.Trace.WriteLine(String.Format("[{0} {1}] {2}",
            T_AddInName,
            String.Format("{0:0#}:{1:0#}:{2:0#}.{3:00#}", DateTime.Now.Hour,
            DateTime.Now.Minute,
            DateTime.Now.Second, DateTime.Now.Millisecond), s_Message));
            System.Diagnostics.Trace.WriteLine("============================");*/

            StringBuilder sb = new StringBuilder(100);
            //sb.AppendLine("============================");
            sb.AppendLine(String.Format("[{0} {1}] {2}",
            "",
            String.Format("{0:0#}:{1:0#}:{2:0#}.{3:00#}", DateTime.Now.Hour,
            DateTime.Now.Minute,
            DateTime.Now.Second, DateTime.Now.Millisecond), s_Message));
            //sb.AppendLine("============================");
            File.AppendAllText(@"C:\Users\admin\Documents\visual studio 2010\Projects\myonemap\myonemap\bin\Debug\log.txt", sb.ToString());
        }

        public static void ShowError(Exception se)
        {
            string text = string.Format("{0}\n{1}", se.Message, se.StackTrace);
            MessageBox.Show(null, text, AssemblyTitle, MessageBoxButtons.OK, MessageBoxIcon.Hand);
        }

        public static string AssemblyTitle
        {
            get
            {
                AssemblyTitleAttribute attribute = (AssemblyTitleAttribute)typeof(Utils).Assembly.GetCustomAttributes(typeof(AssemblyTitleAttribute), true)[0];
                return attribute.Title;
            }
        }


        public static void AddClickHandler(Command command, ICommandEvents_ClickEventHandler clickhHandler)
        {
            EventInfo clickEventInfo = wraper.getEventInfo("ICommandEvents_Event", "Click");
            clickEventInfo.AddEventHandler(command, new ICommandEvents_ClickEventHandler(clickhHandler));
        }

        public static void AddUpdateStateHandler(Command command, ICommandEvents_UpdateStateEventHandler updateStateEventHandler)
        {
            EventInfo updateStateEventInfo = wraper.getEventInfo("ICommandEvents_Event", "UpdateState");
            updateStateEventInfo.AddEventHandler(command, new ICommandEvents_UpdateStateEventHandler(updateStateEventHandler));
        }


        public static List<object> Enumerate<T>(T _object, Type _type)
        {
            var ret = new List<object>();
            if (_type == typeof(Hyperlink))
            {
                IEnumerator enumerator = (_object as IEnumerable).GetEnumerator();
                while (enumerator.MoveNext())
                {
                    ret.Add((Hyperlink)enumerator.Current);
                }
            }
            if (_type == typeof(Event))
            {
                IEnumerator enumerator = (_object as IEnumerable).GetEnumerator();
                while (enumerator.MoveNext())
                {
                    ret.Add((Event)enumerator.Current);
                }
            }
            return ret;
        }

        public static void ReleaseComObject(object _object)
        {
            if (_object != null)
            {
                if (Marshal.IsComObject(_object))
                {
                    try
                    {
                        Marshal.ReleaseComObject(_object);
                        _object = null;
                    }
                    catch (Exception ex)
                    {
                        throw new OneMapException("couldn't release com object", ex);
                    }
                }
            }
        }

        public static void SyncActiveDocumentWithOnenote(ref Mindjet.MindManager.Interop.Application _MindManager)
        {
            if (_MindManager.ActiveDocument == null)
            {
                throw new OneMapException("Couldn't find active document");
            }
            else
            {
                SyncDocumentWithOnenote(_MindManager.ActiveDocument);
            }
        }

        public static void SyncDocumentWithOnenote(Document document)
        {
            SyncForm syncForm = new SyncForm();
            syncForm.Show();
            try
            {
                //todo: don't know what filter is used for
                IEnumerator enumerator = document.Range(MmRange.mmRangeAllTopics, false).GetEnumerator();
                while (enumerator.MoveNext())
                {
                    var topic = enumerator.Current as Topic;
                    if (topic.HasHyperlink)
                    {
                        SyncTopicWithOnenote(topic, ref syncForm);
                        /*topic.Hyperlinks.AsEnumerable()
                            .ToList()
                            .ForEach(x => sb.AppendLine("topic.text " + topic.Text + "-----link: " + x.Address));*/
                    }
                }
            }
            catch (Exception ex)
            {
                throw new OneMapException("Error syncing document",ex);
            }
            finally
            {
                syncForm.Close();
                syncForm.Dispose();
            }
        }

        public static void SyncTopicWithOnenote(Topic topic, ref SyncForm syncForm)
        {
            try
            {
              
                IEnumerator enumerator  = topic.Hyperlinks.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    var hyperlink = enumerator.Current as Hyperlink;
                    if (hyperlink.Address.Contains("onenote:///"))
                    {
                        if (hyperlink.Arguments.Contains("mm-guid"))
                        {
                            SyncHyperLink(ref hyperlink);
                            syncForm.AddListBoxItem(hyperlink.Address);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new OneMapException("Error in SyncTopicWithOnenote ",ex);
            }
        }

        public static void SyncHyperLink(ref Hyperlink onenoteLink)
        {
            string currentHyperLink = onenoteLink.Address  + onenoteLink.Arguments;
            
            try
            {
                string mmguid = currentHyperLink.Substring(currentHyperLink.IndexOf("mm-guid="));
                //mmguid = mmguid.Replace("mm-guid", "").Replace("=", "").Replace("}", "").Replace("{", "").Replace('"', ' ');
                mmguid = mmguid.Replace('"', ' ');
                mmguid = mmguid.Replace(" ", "{", "}", "=", "mm-guid").Trim();
                var newHyperlink = OnenoteUtils.GetHyperLinkBymmGuid(mmguid);
                if (newHyperlink != string.Empty)
                {
                    //todo: does # inside strings adds a bookmark and what should be done.
                    // otherwise it adds the #tail at the end of bookmarks
                    onenoteLink.Bookmark = "";                  
                    newHyperlink += "&mm-guid={" + mmguid + "}";
                    var hl = Utils.GetMindManagerLink(newHyperlink);
                    onenoteLink.Address = hl.Text;
                    onenoteLink.Arguments = hl.Argument;
                }
            }
            catch (OneMapException ex)
            {
                throw new OneMapException("Error in SyncHyperLink",ex);
            }
            catch (Exception ex)
            {
                throw new OneMapException("Error in SyncHyperLink",ex);
            }
            finally
            {

            }
        }

        public static bool IsValidLink( Hyperlink _hyperlink)
        {
            bool retVal = true;

            if (string.IsNullOrEmpty(_hyperlink.Address))
                retVal = false;

            if (string.IsNullOrEmpty(_hyperlink.Arguments))
                retVal = false;

            if (retVal)
            {
                if (!_hyperlink.Address.Contains("onenote"))
                    retVal = false;
            }

            if (retVal)
            {
                if (!_hyperlink.Arguments.Contains("mm-guid"))
                    retVal = false;
            }
            return retVal;
        }

        public static InternalHyperLink GetMindManagerLink(string onenoteLink)
        {
            //if the page has a title then
            if (onenoteLink.Contains("&section-id"))
            {
                int sectionIdIndex = onenoteLink.IndexOf("&section-id");
                int dotOneIndex = onenoteLink.IndexOf(".one#");
                int length = sectionIdIndex - dotOneIndex;
                var title = onenoteLink.Substring(dotOneIndex, length);
                title = title.Replace(".one#", "");
                return new InternalHyperLink(onenoteLink.Substring(0, onenoteLink.IndexOf("&section-id")),
                    onenoteLink.Substring(onenoteLink.IndexOf("&section-id")),title);
            } 
                //when the page does not have a title
            else if (onenoteLink.Contains("#section-id"))
            {
                return new InternalHyperLink(onenoteLink.Substring(0, onenoteLink.IndexOf("#section-id")),
                   onenoteLink.Substring(onenoteLink.IndexOf("#section-id")),string.Empty);
            }
            return new InternalHyperLink(string.Empty,string.Empty,string.Empty);
        }


    }
}