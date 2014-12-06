using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.OneNote;
using myonemap.core;
using myonemap.core.onenote;
using myonemap.enums;
using Application = Microsoft.Office.Interop.OneNote.Application;

namespace myonemap.forms
{
    public partial class SelectOneNotePage : Form
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private static bool exitThread = true;
        public static string retlink = "";
        private Application _onenoteApp;
        private HierarchyType _hierarchyType;

        public SelectOneNotePage(ref Application onteApplication,
                                   HierarchyType hierarchyType)
            : this()
        {
            _onenoteApp = onteApplication;
            _hierarchyType = hierarchyType;
        }

        private SelectOneNotePage()
        {
            exitThread = true;
            retlink = "";
            InitializeComponent();
        }

        private void SelectOneNotePage_Activated(object sender, EventArgs e)
        {
            try
            {
                var t = new Thread(new ThreadStart(QuickDialog));
                t.Start();
                SetOnenoteAsForegroundWindow();

                while (exitThread)
                {
                    Thread.Sleep(50);
                    System.Windows.Forms.Application.DoEvents();
                }
                t.Join();

                SetMindmanagerAsForegroundWindow();
                t.Abort();
            }
            catch (Exception exception)
            {
                throw new OneMapException("error in selectNote->SelectOneNotePage_Activated", exception);
            }
            finally
            {
                this.Close();
            }
        }

        private void QuickDialog()
        {
            IQuickFilingDialog qfDialog;
            qfDialog = _onenoteApp.QuickFiling();
            if (_hierarchyType == HierarchyType.Sections)
            {
                qfDialog.TreeDepth = HierarchyElement.heSections;
                qfDialog.Title = "Import Section to mindmanager";
                qfDialog.AddButton("Import Section", HierarchyElement.heSections, HierarchyElement.heSections, true);
            }
            else if(_hierarchyType == HierarchyType.Pages)
            {
                qfDialog.TreeDepth = HierarchyElement.hePages;
                qfDialog.Title = "select ..... ";
                
                qfDialog.AddButton("Add to mindmanger", HierarchyElement.hePages, HierarchyElement.hePages, true);
            }

            qfDialog.Run(new Callback(ref _onenoteApp,_hierarchyType));
        }

        private void SetOnenoteAsForegroundWindow()
        {
            Process[] p = Process.GetProcessesByName("ONENOTE");
            if (p.Count() > 0)
                SetForegroundWindow(p[0].MainWindowHandle);
        }

        private void SetMindmanagerAsForegroundWindow()
        {
            Process[] p2 = Process.GetProcessesByName("MindManager");
            if (p2.Count() > 0)
                SetForegroundWindow(p2[0].MainWindowHandle);
        }

        private void SelectOneNotePage_Load(object sender, EventArgs e)
        {

        }

        private class Callback : IQuickFilingDialogCallback
        {
            private Application _onenoteApp;
            private HierarchyType _hierarchyType;
            public Callback()
            {
            }

            public Callback(ref Application onenoteApplication,HierarchyType hierarchyType)
            {
                _onenoteApp = onenoteApplication;
                _hierarchyType = hierarchyType;
            }

            public void OnDialogClosed(IQuickFilingDialog qfDialog)
            {
                try
                {
                    if (!string.IsNullOrEmpty(qfDialog.SelectedItem))
                    {
                        if (_hierarchyType == HierarchyType.Pages)
                        {
                            //MessageBox.Show(qfDialog.SelectedItem.ToString());
                            string mmGuid = OnenoteUtils.AddOrGetMeta( qfDialog.SelectedItem,
                                new Meta("mindmanagerguid", Guid.NewGuid().ToString()));
                            string ret = OnenoteUtils.GetLink( qfDialog.SelectedItem);
                            ret += "&mm-guid={" + mmGuid + "}";
                            retlink = ret.ToString();
                        }
                        else if (_hierarchyType == HierarchyType.Sections)
                        {
                            retlink = qfDialog.SelectedItem;
                        }
                    }
                    else
                    {
                        throw new OneMapException("No Item selected");
                    }
                }
                catch (Exception ex)
                {
                    throw new OneMapException("could not add meta->OnDialogClosed+++" + ex.ToStringReflection(), ex, ex.GetBaseException());
                }
                finally
                {
                    exitThread = false;
                }
            }
        }

    }
}
