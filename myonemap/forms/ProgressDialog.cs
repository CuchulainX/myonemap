using System;
using System.Threading;
using System.Windows.Forms;

namespace myonemap.forms
{
    public partial class ProgressDialog : Form
    {
        private string PlayChar = "";
        private string _labelText = "";
        private Thread thread;
        public ProgressDialog()
        {
            InitializeComponent();
            label1.Text = _labelText;
        }

        public int Progress
        {
            get { return pbarProgress.Value; }
            set
            {
                pbarProgress.Value = value;
                //Application.DoEvents();
            }
        }

        public string Message
        {
            set
            {
                label1.Text = value;
                _labelText = value;
            }
        }

        public void SetPbarVisiblity(bool visible)
        {
            pbarProgress.Visible = visible;
        }

        public void PlayCursor()
        {
            thread = new Thread(new ThreadStart(ShowCursor));
            thread.Start();
        }

        public void StopCursor()
        {
            if (thread.ThreadState == ThreadState.Running)
            {
                thread.Abort();
                thread = null;
            }
        }

        private void ShowCursor()
        {
            while (true)
            {
                
                switch (PlayChar)
                {
                    case "":
                        PlayChar = "|";
                        break;
                    case "|":
                        PlayChar = "/";
                        break;
                    case "/":
                        PlayChar = "-";
                        break;
                    case "-":
                        PlayChar=@"\";
                        break;
                    case @"\":
                        PlayChar = "|";
                        break;
                }
                label1.Text = _labelText + "   " + PlayChar;
                Thread.Sleep(300);
                Application.DoEvents();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {

        }

    }
}
