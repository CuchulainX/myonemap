using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace myonemap.core.onenote
{
    internal class HResultInfo
    {
        public int HrCode;
        public string HrDesc;
        public string HrName;
        public HResultInfo(string hrName, uint hrCode, string hrDesc)
        {
            HrName = hrName;
            HrCode = (int)hrCode;
            HrDesc = hrDesc;
        }
    }
    internal static class OneNoteHresultDescriptions
    {
        private static List<HResultInfo> col = new List<HResultInfo>();
        static OneNoteHresultDescriptions()
        {
            col.Add(new HResultInfo("hrMalformedXML", 0x80042000, "The XML is not well-formed."));
            col.Add(new HResultInfo("hrInvalidXML", 0x80042001, "The XML is invalid."));
            col.Add(new HResultInfo("hrCreatingSection", 0x80042002, "The section could not be created."));
            col.Add(new HResultInfo("hrOpeningSection", 0x80042003, "The section could not be opened."));
            col.Add(new HResultInfo("hrSectionDoesNotExist", 0x80042004, "The section does not exist."));
            col.Add(new HResultInfo("hrPageDoesNotExist", 0x80042005, "The page does not exist."));
            col.Add(new HResultInfo("hrFileDoesNotExist", 0x80042006, "The file does not exist."));
            col.Add(new HResultInfo("hrInsertingImage", 0x80042007, "The image could not be inserted."));
            col.Add(new HResultInfo("hrInsertingInk", 0x80042008, "The ink could not be inserted."));
            col.Add(new HResultInfo("hrInsertingHtml", 0x80042009, "The HTML could not be inserted."));
            col.Add(new HResultInfo("hrNavigatingToPage", 0x8004200a, "The page could not be opened."));
            col.Add(new HResultInfo("hrSectionReadOnly", 0x8004200b, "The section is read-only."));
            col.Add(new HResultInfo("hrPageReadOnly", 0x8004200c, "The page is read-only."));
            col.Add(new HResultInfo("hrInsertingOutlineText", 0x8004200d, "The outline text could not be inserted."));
            col.Add(new HResultInfo("hrPageObjectDoesNotExist", 0x8004200e, "The page object does not exist."));
            col.Add(new HResultInfo("hrBinaryObjectDoesNotExist", 0x8004200f, "The binary object does not exist."));
            col.Add(new HResultInfo("hrLastModifiedDateDidNotMatch", 0x80042010, "The last modified date does not match."));
            col.Add(new HResultInfo("hrGroupDoesNotExist", 0x80042011, "The section group does not exist."));
            col.Add(new HResultInfo("hrPageDoesNotExistInGroup", 0x80042012, "The page does not exist in the section group."));
            col.Add(new HResultInfo("hrNoActiveSelection", 0x80042013, "There is no active selection."));
            col.Add(new HResultInfo("hrObjectDoesNotExist", 0x80042014, "The object does not exist."));
            col.Add(new HResultInfo("hrNotebookDoesNotExist", 0x80042015, "The notebook does not exist."));
            col.Add(new HResultInfo("hrInsertingFile", 0x80042016, "The file could not be inserted."));
            col.Add(new HResultInfo("hrInvalidName", 0x80042017, "The name is invalid."));
            col.Add(new HResultInfo("hrFolderDoesNotExist", 0x80042018, "The folder (section group) does not exist."));
            col.Add(new HResultInfo("hrInvalidQuery", 0x80042019, "The query is invalid."));
            col.Add(new HResultInfo("hrFileAlreadyExists", 0x8004201a, "The file already exists."));
            col.Add(new HResultInfo("hrSectionEncryptedAndLocked", 0x8004201b, "The section is encrypted and locked."));
            col.Add(new HResultInfo("hrDisabledByPolicy", 0x8004201c, "The action is disabled by a policy."));
            col.Add(new HResultInfo("hrNotYetSynchronized", 0x8004201d, "OneNote has not yet synchronized content."));
            col.Add(new HResultInfo("hrLegacySection", 0x8004201E, "The section is from OneNote 2007 or earlier."));
            col.Add(new HResultInfo("hrMergeFailed", 0x8004201F, "The merge operation failed."));
            col.Add(new HResultInfo("hrInvalidXMLSchema", 0x80042020, "The XML Schema is invalid."));
            col.Add(new HResultInfo("hrFutureContentLoss", 0x80042022, "Content loss has occurred (from future versions of OneNote)."));
            col.Add(new HResultInfo("hrTimeOut", 0x80042023, "The action timed out."));
            col.Add(new HResultInfo("hrRecordingInProgress", 0x80042024, "Audio recording is in progress."));
            col.Add(new HResultInfo("hrUnknownLinkedNoteState", 0x80042025, "The linked-note state is unknown."));
            col.Add(new HResultInfo("hrNoShortNameForLinkedNote", 0x80042026, "No short name exists for the linked note."));
            col.Add(new HResultInfo("hrNoFriendlyNameForLinkedNote", 0x80042027, "No friendly name exists for the linked note."));
            col.Add(new HResultInfo("hrInvalidLinkedNoteUri", 0x80042028, "The linked note URI is invalid."));
            col.Add(new HResultInfo("hrInvalidLinkedNoteThumbnail", 0x80042029, "The linked note thumbnail is invalid."));
            col.Add(new HResultInfo("hrImportLNTThumbnailFailed", 0x8004202A, "The importation of linked note thumbnail failed."));
            col.Add(new HResultInfo("hrUnreadDisabledForNotebook", 0x8004202B, "Unread highlighting is disabled for the notebook."));
            col.Add(new HResultInfo("hrInvalidSelection", 0x8004202C, "The selection is invalid."));
            col.Add(new HResultInfo("hrConvertFailed", 0x8004202D, "The conversion failed."));
            col.Add(new HResultInfo("hrRecycleBinEditFailed", 0x8004202E, "Edit failed in the Recycle Bin."));
        }

        /*public static string GetErrorDescription(uint errorCode)
        {
            HResultInfo hr = col.FirstOrDefault(x => x.HrCode == errorCode);
            if (hr != null)
                return hr.HrDesc;
            return string.Empty;
        }
*/
        public static string GetOnenoteErrorDescription(int errorCode)
        {
            HResultInfo hr = col.FirstOrDefault(x => x.HrCode == errorCode);
            if (hr != null)
                return hr.HrDesc;
            return string.Empty;
        }

        public static string GetErrorDescription(int errorCode)
        {
            var ret = string.Empty;
            ret = GetOnenoteErrorDescription(errorCode);
            if (string.IsNullOrEmpty(ret))
            {
                ret = Marshal.GetExceptionForHR(errorCode).Message;
            }
            return ret;
        }
    }
}
