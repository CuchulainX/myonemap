namespace myonemap.core.onenote
{
    public class Meta
    {
        public Meta(string name, string content)
        {
            Name = name;
            Content = content;
        }

        public string Name { get; set; }
        public string Content { get; set; }
    }
}
