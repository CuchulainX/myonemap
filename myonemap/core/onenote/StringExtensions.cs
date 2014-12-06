namespace myonemap.core.onenote
{
    public static class StringExtensions
    {
        public static string Replace(this string source,string replaceString, params string[] replaceStrings)
        {
            foreach (var item in replaceStrings)
            {
                source = source.Replace(item, replaceString);
            }
            return source;
        }
    }
}
