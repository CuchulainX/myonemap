using System;

namespace myonemap
{
    public class SectionImport
    {
        public delegate void SectionImportEvent(object sender, SectionImportEventArgs e);

        public event SectionImportEvent SectionImport1;

        public void RaiseEvent(string sectionId)
        {
            if (SectionImport1 != null)
            {
                //SectionImport(this,sectionId)
            }
        }
    }

    public class SectionImportEventArgs :EventArgs
    {
        public SectionImportEventArgs()
        {
        }
    }
}