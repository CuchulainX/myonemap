using System;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Mindjet.MindManager.Interop;
using myonemap.core.mindmanager;
using myonemap.core.onenote;
using myonemap.enums;
using myonemap.forms;
using myonemap.Properties;
using Control = Mindjet.MindManager.Interop.Control;
using MindmanagerApplication = Mindjet.MindManager.Interop.Application;
using Utilities = Mindjet.MindManager.Interop.Utilities;

// icon from "http://www.flaticon.com" 
namespace myonemap
{

    #region Read me for Add-in installation and setup information.
    // When run, the Add-in wizard prepared the registry for the Add-in.
    // At a later time, if the Add-in becomes unavailable for reasons such as:
    //   1) You moved this project to a computer other than which is was originally created on.
    //   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
    //   3) Registry corruption.
    // you will need to re-register the Add-in by building the myonemapSetup project, 
    // right click the project in the Solution Explorer, then choose install.
    #endregion

    /// <summary>
    ///   The object for implementing an Add-in.
    /// </summary>
    /// <seealso class='IDTExtensibility2' />
    [GuidAttribute("B7AA323D-F6E7-414A-8CBA-DD5F28E607D2"), ProgId("myonemap.Connect"), ComVisible(true), CLSCompliant(false)]
    public partial class Connect : Extensibility.IDTExtensibility2
    {
        /// <summary>
        ///		Implements the constructor for the Add-in object.
        ///		Place your initialization code within this method.
        /// </summary>
        public Connect()
        {
            AppDomain.CurrentDomain.UnhandledException +=
                new UnhandledExceptionEventHandler(LastChanceHandler);
        }
        static void LastChanceHandler(object sender, UnhandledExceptionEventArgs args)
        {
            try
            {
                Exception e = (Exception)args.ExceptionObject;

                //  Console.WriteLine("Unhandled exception == " + e.ToString());
                if (args.IsTerminating)
                {
                    //    Console.WriteLine("The application is terminating");
                }
                else
                {
                    //  Console.WriteLine("The application is not terminating");
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine("Inside Last Chance Handler: " + e.Message);
            }
            finally
            {
                // Add other exception logging or cleanup code here.
            }
        }
        /// <summary>
        ///      Implements the OnConnection method of the IDTExtensibility2 interface.
        ///      Receives notification that the Add-in is being loaded.
        /// </summary>
        /// <param term='application'>
        ///      Root object of the host application.
        /// </param>
        /// <param term='connectMode'>
        ///      Describes how the Add-in is being loaded.
        /// </param>
        /// <param term='addInInst'>
        ///      Object representing this Add-in.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode,
                                object addInInst, 
                                ref System.Array custom)
        {
            _MindManager = (Mindjet.MindManager.Interop.Application)application;
            addInInstance = addInInst;
            
            AddCommands();
        }

        private void AddCommands()
        {
            try
            {
                #region commands

                this._HeaderCommand = _MindManager.Commands.Add("myonemap.Connect", "onenote_tools_header");
                _HeaderCommand.ToolTip = "Onenote Tools";
                _HeaderCommand.Caption = "Onenote Tools";
                var path = Path.GetTempPath() + "header.png";
                Properties.Resources.header.Save(path,ImageFormat.Png);
                _HeaderCommand.LargeImagePath = path;

                //Utils.AddClickHandler(_HeaderCommand, HeaderClick);
                Utils.AddUpdateStateHandler(_HeaderCommand, HeaderUpdateState);

                //*******    Add Link to topic
                this._AddLinkToSelectedTopicCommand = _MindManager.Commands.Add("myonemap.Connect", "add_link");
                _AddLinkToSelectedTopicCommand.ToolTip = "Add Link To Topic";
                _AddLinkToSelectedTopicCommand.Caption = "Add Link To Topic";
                Utils.AddClickHandler(_AddLinkToSelectedTopicCommand, AddLinkToTopic);
                Utils.AddUpdateStateHandler(_AddLinkToSelectedTopicCommand, AddLinkToTopicUpdateState);

                //******  Sync Active topic with onenote
                _SyncActiveTopicWithOnenote = _MindManager.Commands.Add("myonemap.Connect", "sync_topic");
                _SyncActiveTopicWithOnenote.ToolTip = "Synchronize topic links with onenote";
                _SyncActiveTopicWithOnenote.Caption = "Synchronize topic links with onenote";
                Utils.AddClickHandler(_SyncActiveTopicWithOnenote, SyncTopicLinksWithOnnote);
                Utils.AddUpdateStateHandler(_SyncActiveTopicWithOnenote, SyncTopicLinkWithOnenoteUpdateState);

                //******** Sync Active Document With onenote
                _SyncActiveDocumentWithOnenote = _MindManager.Commands.Add("myonemap.Connect", "sync_doc");
                _SyncActiveDocumentWithOnenote.ToolTip = "Synchronize document links with onenote";
                _SyncActiveDocumentWithOnenote.Caption = "Synchronize document with onenote";
                Utils.AddClickHandler(_SyncActiveDocumentWithOnenote, SyncDocumentWithOnenote);
                Utils.AddUpdateStateHandler(_SyncActiveDocumentWithOnenote, SyncDocumentWithOnenoteUpdateState);


                //*********** import a section 
                _ImportCurrentSection = _MindManager.Commands.Add("myonemap.Connect", "import_section");
                _ImportCurrentSection.ToolTip = "Import Section";
                _ImportCurrentSection.Caption = "Import Section";
                Utils.AddClickHandler(_ImportCurrentSection, ImportCurrentOnenoteSection);
                Utils.AddUpdateStateHandler(_ImportCurrentSection, ImportCurrentOnenoteSectionUpdateState);

                //*************** Managing onenote links
                // todo : implement 
                /*
                _ManageOnenoteLinks = _MindManager.Commands.Add("myonemap.Connect", "manage_links");
                _ManageOnenoteLinks.ToolTip = "Manage onenote links for this document";
                _ManageOnenoteLinks.Caption = "Manage onenote links";
                Utils.AddClickHandler(_ManageOnenoteLinks, MangeOnenoteLinks);
                Utils.AddUpdateStateHandler(_ManageOnenoteLinks, ManageOnenoteLinksUpdateState);
                */

                #endregion

                _HomeTab = Utils.GetRibbonTabByName(ref _MindManager, "Home");

                ribbonGroup = _HomeTab.Groups.Add(0, "myonemap", "http://www.myonemap.com/myonemap", "");
                Control button = ribbonGroup.GroupControls.AddButton(_HeaderCommand, 0);
                //button.Controls.AddButton(_HeaderCommand, 0);
                button.Controls.AddButton(_AddLinkToSelectedTopicCommand, 0);
                button.Controls.AddButton(_SyncActiveTopicWithOnenote, 0);
                button.Controls.AddButton(_SyncActiveDocumentWithOnenote, 0);
                //button.Controls.AddButton(_AddNoteToCurrentSection, 0);
                button.Controls.AddButton(_ImportCurrentSection, 0);

                //todo: implement
                //button.Controls.AddButton(_ManageOnenoteLinks, 0);
                
            }
            catch (Exception ex)
            {
                Utils.ShowError(ex);
            }
            
        }

        #region updteStates
        private void SyncTopicLinkWithOnenoteUpdateState(ref bool penabled, ref bool pchecked)
        {
            penabled = true;
            //this is not efficient
/*
            if (!OnenoteUtils.WSearchIsOn())
                penabled = false;*/
        }

        private void SyncDocumentWithOnenoteUpdateState(ref bool penabled, ref bool pchecked)
        {
            penabled = true;
         /*   if (!OnenoteUtils.WSearchIsOn())
                penabled = false;*/
        }



        private void ImportCurrentOnenoteSectionUpdateState(ref bool penabled, ref bool pchecked)
        {
            penabled = true;
            // when there's not document 
            if (_MindManager.ActiveDocument == null)
            {
                penabled = false;
            }
            //when no topic is selected
            else
            {
                if (_MindManager.ActiveDocument.Selection.PrimaryTopic == null)
                {
                    penabled = false;
                }
            }
        }

       
        public void AddLinkToTopicUpdateState(ref bool enabled, ref bool chacked)
        {
            enabled = true;
            // when there's not document 
            if (_MindManager.ActiveDocument == null)
            {
                enabled = false;
            }
            //when no topic is selected
            else
            {
                if (_MindManager.ActiveDocument.Selection.PrimaryTopic == null)
                {
                    enabled = false;
                }
            }
        }
        #endregion

        private void SyncTopicLinksWithOnnote()
        {
            if (OnenoteUtils.WSearchIsOn())
            {
                var seltopic = _MindManager.ActiveDocument.Selection.PrimaryTopic;
                SyncForm syncForm = new SyncForm();
                try
                {
                    Utils.SyncTopicWithOnenote(seltopic, ref syncForm);

                }
                catch (Exception ex)
                {
                    throw new OneMapException("SyncDocumentWithOneNote-> couldn't synchronize document with onenote", ex);
                }
                finally
                {
                    syncForm.Close();
                }
            }
            else
                MessageBox.Show("WSearch is not on!!!!");
        }

        public void SyncDocumentWithOnenote()
        {
            if (OnenoteUtils.WSearchIsOn())
            {
                try
                {
                    Utils.SyncActiveDocumentWithOnenote(ref _MindManager);
                }
                catch (Exception ex)
                {
                    throw new OneMapException("SyncDocumentWithOneNote-> couldn't synchronize document with onenote", ex);
                }
              //  _MindManager.ActiveDocument.Save();
            }
            else
                MessageBox.Show("WSearch is not on!!!!");
        }

        // a link manager form maybe later
       /* public void MangeOnenoteLinks()
        {
            ManageLinksfrm manageLinksfrm = new ManageLinksfrm(ref _MindManager);
            try
            {
                manageLinksfrm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening dialog" + ex.Message);
            }
            finally
            {
                manageLinksfrm.Close();
                manageLinksfrm.Dispose();
            }
        }*/

        public void ImportCurrentOnenoteSection()
        {
            ImportDialog importDialog = new ImportDialog( HierarchyType.Sections, ref _MindManager);

            try
            {
                importDialog.QuickDialog();
            }
            catch (OneMapException exception)
            {

            }
            catch (Exception ex)
            {

            }
        }

        public void HeaderUpdateState(ref bool enabled, ref bool chacked)
        {
            enabled = true;
        }

        private void AddLinkToTopic()
        {
            ImportDialog importDialog = new ImportDialog( HierarchyType.Pages, ref _MindManager);
            
            try
            {
                importDialog.QuickDialog();
            }
            catch (Exception ex)
            {

            }
           
        }
     

        public void HeaderClick()
        {
            
        }


        /// <summary>
        ///     Implements the OnDisconnection method of the IDTExtensibility2 interface.
        ///     Receives notification that the Add-in is being unloaded.
        /// </summary>
        /// <param term='disconnectMode'>
        ///      Describes how the Add-in is being unloaded.
        /// </param>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
        {
            //Utils.Pt(String.Format(".Connect.OnDisconnection(disconnectMode = {0})",
//disconnectMode.ToString()));

            //===========================================================
            //DumpInfo();
            //=======================================================
            
            /*addInInstance = null;
            _HeaderCommand = null;
            _AddLinkToSelectedTopicCommand = null;
            _HomeTab = null;
            ribbonGroup = null;
            _MindManager = null;
            _SyncActiveTopicWithOnenote = null;
            _SyncActiveTopicWithOnenote = null;
            _SyncActiveDocumentWithOnenote = null;
            _ImportCurrentSection = null;
            _ManageOnenoteLinks = null;*/
            OneNote.Dispose();
            Utils.ReleaseComObject(_HeaderCommand);
            Utils.ReleaseComObject(_AddLinkToSelectedTopicCommand);
            Utils.ReleaseComObject(_HomeTab);
            Utils.ReleaseComObject(ribbonGroup);
            Utils.ReleaseComObject(_SyncActiveTopicWithOnenote);
            Utils.ReleaseComObject(_SyncActiveTopicWithOnenote);
            Utils.ReleaseComObject(_SyncActiveDocumentWithOnenote);
            Utils.ReleaseComObject(_ImportCurrentSection);
            Utils.ReleaseComObject(_MindManager);
        }

        /// <summary>
        ///      Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
        ///      Receives notification that the collection of Add-ins has changed.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnAddInsUpdate(ref System.Array custom)
        {
          //  Utils.Pt(".Connect.OnStartupComplete()");
        }

        /// <summary>
        ///      Implements the OnStartupComplete method of the IDTExtensibility2 interface.
        ///      Receives notification that the host application has completed loading.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnStartupComplete(ref System.Array custom)
        {

        }

        /// <summary>
        ///      Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
        ///      Receives notification that the host application is being unloaded.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnBeginShutdown(ref System.Array custom)
        {
           // Utils.Pt(".Connect.OnBeginShutdown()");
        }


        private MindmanagerApplication _MindManager;
        private object addInInstance;
        private Command _HeaderCommand;
        private Command _AddLinkToSelectedTopicCommand;
        private Command _SyncActiveTopicWithOnenote;
        private Command _SyncActiveDocumentWithOnenote;

        private Command _ImportCurrentSection;
        private Command _ManageOnenoteLinks;

        private ribbonTab _HomeTab;
        private RibbonGroup ribbonGroup;

    }
}