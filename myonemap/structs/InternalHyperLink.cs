namespace myonemap.structs
{
    public struct InternalHyperLink
    {
        public string Text;
        public string Argument;
        public bool HasTitle;
        public string Title;

        public InternalHyperLink(string text, string argument,string title)
        {
            if (string.IsNullOrEmpty(title))
            {
                HasTitle = false;
                Title = string.Empty;
            }
            else
            {
                HasTitle = true;
                Title = title.Replace("%20", " ");
            }
            Text = text;
            Argument = argument;
        }
    }
}