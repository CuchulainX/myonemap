using System.Windows.Forms;

namespace myonemap.forms
{
    public partial class SyncForm : Form
    {
        public SyncForm()
        {
            InitializeComponent();
        }

        public void AddListBoxItem(string item)
        {
            listBox1.Items.Add(item);
            Application.DoEvents();
        }
    }
}
