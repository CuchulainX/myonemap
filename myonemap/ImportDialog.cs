using System;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.OneNote;
using Mindjet.MindManager.Interop;
using myonemap.core;
using myonemap.core.mindmanager;
using myonemap.core.onenote;
using myonemap.enums;
using myonemap.forms;
using Application = System.Windows.Forms.Application;
using onenoteApplication = Microsoft.Office.Interop.OneNote.Application;
using MindmanagerApplication = Mindjet.MindManager.Interop.Application;

namespace myonemap
{
    public class ImportDialog
    {
        private MindmanagerApplication _mindManager;
        private HierarchyType _hierarchyType;
        private ulong _windowHandle;


        public ImportDialog(HierarchyType hierarchyType,
                            ref MindmanagerApplication mindManager)
        {
            _hierarchyType = hierarchyType;
            _mindManager = mindManager;
            _windowHandle = (ulong)mindManager.hWnd;
        }

        public void QuickDialog()
        {
            IQuickFilingDialog qfDialog;
            qfDialog = OneNote.Instance().QuickFiling();

            if (_hierarchyType == HierarchyType.Sections)
            {
                qfDialog.TreeDepth = HierarchyElement.heSections;
                qfDialog.Title = "Import Section to mindmanager";
                qfDialog.ParentWindowHandle = _windowHandle;
                qfDialog.AddButton("Import Section", HierarchyElement.heSections, HierarchyElement.heSections, true);
            }
            else if (_hierarchyType == HierarchyType.Pages)
            {
                qfDialog.TreeDepth = HierarchyElement.hePages;
                qfDialog.Title = "Select Page ";
                qfDialog.ParentWindowHandle = _windowHandle;
                qfDialog.AddButton("Add page Link to Topic", HierarchyElement.hePages, HierarchyElement.hePages, true);
            }

            qfDialog.Run(new Callback( _hierarchyType, ref _mindManager));
        }


        private class Callback : IQuickFilingDialogCallback
        {
            private HierarchyType _hierarchyType;
            private MindmanagerApplication _mindManager;
            public Callback()
            {
            }

            public Callback(HierarchyType hierarchyType,
                            ref MindmanagerApplication mindManager)
            {
                _hierarchyType = hierarchyType;
                _mindManager = mindManager;
            }

            public void OnDialogClosed(IQuickFilingDialog qfDialog)
            {
                try
                {
                    // On cancel the IQuickFilingDialog.pressedButton returns -1
                    // other buttons start from 0 to ...
                    if (qfDialog.PressedButton == 0)
                    {
                        if (_hierarchyType == HierarchyType.Pages)
                        {
                            ImportPageLink(qfDialog.SelectedItem);
                        }
                        else if (_hierarchyType == HierarchyType.Sections)
                        {
                            ImportSectionLinks(qfDialog.SelectedItem);
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new OneMapException(">OnDialogClosed+++" + ex.ToStringReflection(), ex, ex.GetBaseException());
                }
            }

            private void ImportSectionLinks(string selectedItem)
            {
                var retstring = selectedItem;
                var pDialog = new ProgressDialog();

                try
                {
                    var owner = new Win32Window(Process.GetCurrentProcess().MainWindowHandle);
                    pDialog.Show(owner);
                    var selTopic = _mindManager.ActiveDocument.Selection.PrimaryTopic;
                    if (selTopic != null)
                    {
                        var pageIds = OnenoteUtils.GetSectionPageIds(retstring);
                        int count = pageIds.Count();
                        int block = 100 / count;
                        pDialog.Progress = 100 % count;
                        string mmLink;
                        foreach (var pageId in pageIds)
                        {
                            mmLink = OnenoteUtils.GetMindManagerLink(pageId);
                            var hl = Utils.GetMindManagerLink(mmLink);

                            pDialog.Progress += block;
                            pDialog.Message = hl.Title;
                            Topic topic = selTopic.AddSubTopic(hl.Title);
                            if (!string.IsNullOrEmpty(mmLink))
                            {
                                Hyperlink hyperlink = topic.Hyperlinks.AddHyperlink(hl.Text);
                                hyperlink.Arguments = hl.Argument;
                            }

                            Application.DoEvents();
                            //_MindManager.ActiveDocument.Save();
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

            private void ImportPageLink(string selectedItem)
            {
                string retstring = OnenoteUtils.GetMindManagerLink(selectedItem);

                if (!string.IsNullOrEmpty(retstring))
                {
                    var seltopic = _mindManager.ActiveDocument.Selection.PrimaryTopic;
                    var hl = Utils.GetMindManagerLink(retstring);
                    Hyperlink hyperlink = seltopic.Hyperlinks.AddHyperlink(hl.Text);
                    hyperlink.Arguments = hl.Argument;

                    // _mindManager.ActiveDocument.Save();
                }
            }

        }
    }
}