using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Mindjet.MindManager.Interop;
using myonemap.core;
using myonemap.core.mindmanager;
using myonemap.core.onenote;
using Application = Mindjet.MindManager.Interop.Application;
using Color = System.Drawing.Color;

namespace myonemap.forms
{
    public partial class ManageLinksfrm : Form
    {
        private Application _app = null;
        public ManageLinksfrm(ref Application app)
        {
            InitializeComponent();
            _app = app;
        }


        private void ManageLinksfrm_Load(object sender, EventArgs e)
        {
            Initialize();
        }

        private void Initialize()
        {
            lstLinks.Items.Clear();
            var document = _app.ActiveDocument;
            var linkList = new List<Tuple<string, bool, string,string>>();
            try
            {
                //todo: don't know what filter is used for
                IEnumerator enumerator = document.Range(MmRange.mmRangeAllTopics, false).GetEnumerator();
                while (enumerator.MoveNext())
                {
                    var topic = enumerator.Current as Topic;
                    if (topic != null && topic.HasHyperlink)
                    {
                        foreach (var hyperlink in topic.Hyperlinks.AsEnumerable())
                        {
                            if (Utils.IsValidLink(hyperlink))
                            {
                                string currentHyperLink = hyperlink.Address + "#" + hyperlink.Bookmark + hyperlink.Arguments;
                                string mmguid = currentHyperLink.Substring(currentHyperLink.IndexOf("mm-guid="));
                                mmguid =
                                    mmguid.Replace("mm-guid", "")
                                        .Replace("=", "")
                                        .Replace("}", "")
                                        .Replace("{", "")
                                        .Replace('"', ' ');
                                var newHyperlink = OnenoteUtils.GetHyperLinkBymmGuid(mmguid);
                                // MessageBox.Show("newhyperlink: " + newHyperlink);
                               
                                    //if newHyperlink is not empty then onenote page exists
                                    if (!string.IsNullOrEmpty(newHyperlink))
                                    {
                                        var foundHyperLink = linkList.Any(x => x.Item4.Contains(newHyperlink));
                                        if (!foundHyperLink)
                                        {
                                            if (currentHyperLink.Contains("&section-id"))
                                            {
                                                linkList.Add(
                                                    new Tuple<string, bool, string, string>(
                                                        currentHyperLink.Substring(0,
                                                            currentHyperLink.IndexOf("&section-id")),
                                                        true, topic.Guid, newHyperlink));
                                            }
                                            else if (currentHyperLink.Contains("section-id"))
                                            {
                                                linkList.Add(
                                                    new Tuple<string, bool, string, string>(
                                                        currentHyperLink.Substring(0,
                                                            currentHyperLink.IndexOf("section-id")),
                                                        false, topic.Guid, newHyperlink));
                                            }
                                        }
                                    }
                                    else
                                    {
                                            linkList.Add(
                                                new Tuple<string, bool, string, string>(
                                                    currentHyperLink.Substring(0,
                                                        currentHyperLink.IndexOf("&section-id")), false,
                                                    topic.Guid, newHyperlink));
                                    }
                            }
                        }
                    }
                    Utils.ReleaseComObject(topic);
                }
                //LinkList.ForEach(x => lstLinks.Items.Add(x.Address));
                linkList.ForEach(x =>
                {
                    if (!x.Item2)
                    {
                        int index = lstLinks.Items.Add(new MyListBoxItem(Color.Red, x.Item1));
                        
                    }
                    else
                    {
                        lstLinks.Items.Add(new MyListBoxItem(Color.Blue, x.Item1));
                    }
                });
            }
            catch (Exception ex)
            {
                throw new OneMapException("ManageLinksfrm_Load()->Error Managing links", ex);
            }
            finally
            {
                Utils.ReleaseComObject(document);
            }
        }

        private void lstLinks_DrawItem(object sender, DrawItemEventArgs e)
        {
            var item = lstLinks.Items[e.Index] as MyListBoxItem; // Get the current item and cast it to MyListBoxItem
            if (item != null)
            {
                e.Graphics.DrawString( // Draw the appropriate text in the ListBox
                    item.Message, // The message linked to the item
                    lstLinks.Font, // Take the font from the listbox
                    new SolidBrush(item.ItemColor), // Set the color 
                    0, // X pixel coordinate
                    e.Index * lstLinks.ItemHeight // Y pixel coordinate.  Multiply the index by the ItemHeight defined in the listbox.
                );
            }
            else
            {
                // The item isn't a MyListBoxItem, do something about it
            }
        }

        private void btnSync_Click(object sender, EventArgs e)
        {
            try
            {
                Utils.SyncActiveDocumentWithOnenote(ref _app);
                _app.ActiveDocument.Save();
                Initialize();
            }
            catch (Exception ex)
            {
                throw new OneMapException("btnSync_Click-> couldn't synchronize document with onenote", ex);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            var document = _app.ActiveDocument;
            try
            {
                IEnumerator enumerator = document.Range(MmRange.mmRangeAllTopics, false).GetEnumerator();
                while (enumerator.MoveNext())
                {
                    var topic = enumerator.Current as Topic;
                    if (topic != null && topic.HasHyperlink)
                    {
                        foreach (var hyperlink in topic.Hyperlinks.AsEnumerable())
                        {
                            if (Utils.IsValidLink(hyperlink))
                            {
                                string currentHyperLink = hyperlink.Address + "#" + hyperlink.Bookmark +
                                                          hyperlink.Arguments;
                                string mmguid = currentHyperLink.Substring(currentHyperLink.IndexOf("mm-guid="));
                                mmguid =
                                    mmguid.Replace("mm-guid", "")
                                        .Replace("=", "")
                                        .Replace("}", "")
                                        .Replace("{", "")
                                        .Replace('"', ' ');
                                var newHyperlink = OnenoteUtils.GetHyperLinkBymmGuid(mmguid);
                                // MessageBox.Show("newhyperlink: " + newHyperlink);
                                if (string.IsNullOrEmpty(newHyperlink))
                                {
                                    hyperlink.Bookmark = "";
                                    hyperlink.Arguments = "";
                                    hyperlink.Delete();
                                }
                            }
                        }
                    }
                }
                Initialize();
            }
            catch (COMException cex)
            {
                throw new OneMapException("btnDelete_Click-> couldn't delete link",cex.InnerException,cex.GetBaseException());
            }
            catch (Exception ex)
            {
                throw new OneMapException("btnDelete_Click-> couldn't delete link");
            }
            finally
            {
                Utils.ReleaseComObject(document);
            }
        }
    }

    public class MyListBoxItem
    {
        public MyListBoxItem(System.Drawing.Color c, string m)
        {
            ItemColor = c;
            Message = m;
        }
        public System.Drawing.Color ItemColor { get; set; }
        public string Message { get; set; }
    }
}
