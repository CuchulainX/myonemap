using System;
using System.Windows.Forms;

namespace myonemap.core
{
    class Win32Window : IWin32Window
    {
        private readonly IntPtr _hwnd;
        public Win32Window(IntPtr handle)
        {
            _hwnd = handle;
        }
        public IntPtr Handle
        {
            get
            {
                return _hwnd;
            }
        }
    }
}