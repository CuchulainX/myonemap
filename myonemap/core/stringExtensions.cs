using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using Mindjet.MindManager.Interop;

namespace myonemap.core
{
    public static class StringExtensions
    {
        public static string Prepend(this string source, string head)
        {
            return head + source + "\n";
        }

        public static void AddToSb(this string source, StringBuilder sb)
        {
            sb.Append(source);
        }

        public static IEnumerable<IRibbonTab> AsEnumerable(this RibbonTabs source)
        {
            IEnumerator enumerator = source.GetEnumerator();
            while (enumerator.MoveNext())
            {
                yield return (enumerator.Current as IRibbonTab);
            }
        }

        public static IEnumerable<RibbonGroup> AsEnumerable(this RibbonGroups sourceGroups)
        {
            IEnumerator enumerator = sourceGroups.GetEnumerator();
            while (enumerator.MoveNext())
            {
                yield return (enumerator.Current as RibbonGroup);
            }
        }

        public static IEnumerable<Topic> AsEnumerable(this Topics  source)
        {
            IEnumerator enumerator = source.GetEnumerator();
            while (enumerator.MoveNext())
            {
                yield return (enumerator.Current as Topic);
            }
        }

        public static IEnumerable<Hyperlink> AsEnumerable(this Hyperlinks source)
        {
            IEnumerator enumerator = source.GetEnumerator();
            while (enumerator.MoveNext())
            {
                yield return (enumerator.Current as Hyperlink);
            }
        }

        
        public static string ToStr(this object source)
        {
            string ret = "";
            if (source is List<RibbonGroup>)
            {
                foreach (var ribbonGroup in (source as List<RibbonGroup>))
                {
                    ret += ribbonGroup.DisplayName + "-";
                }
            }
            return ret;
        }

        public static IEnumerable<Control> AsEnumerable(this Controls source)
        {
            IEnumerator enumerator = source.GetEnumerator();
            while (enumerator.MoveNext())
            {
                yield return (enumerator.Current as Control);
            }
        }


        public static string ToStringReflection<T>(this T @this)
        {
            //It takes thosePropertyInfo objectrs and uses the Name Property
            // and GetValue() method to create a descriptive string for each
            // object.the result of the linq query are then passed into Join()
            //  on the string class, which takes easch descriptive string and
            // joins them up, seperating them with || separator.
            return string.Join(Constants.Separator,
                new List<string>(
                    from prop in @this.GetType().GetProperties(
                        BindingFlags.Instance | BindingFlags.Public)
                    where prop.CanRead
                    select string.Format("{0}: {1}", prop.Name, prop.GetValue(@this, null))).ToArray());
        }

    }

    public static class Constants
    {
        public const string Separator = " || ";
    }
}
    
    
