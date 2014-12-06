using Microsoft.Office.Interop.OneNote;
using myonemap.core.mindmanager;

namespace myonemap.core.onenote
{
    public class OneNote
    {
        private static Application _onenoteApp=null;
        public static Application Instance()
        {
            if (_onenoteApp == null)
            {
                _onenoteApp = new Application();
            }
            return _onenoteApp;
        }

        public static void Dispose()
        {
            Utils.ReleaseComObject(_onenoteApp);
        }
    }

}
