using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using Mindjet.MindManager.Interop;

namespace Utilities
{
    public static class wraper
    {
        public static EventInfo getEventInfo(Type t, string eventName)
        {
            if (typeof(ICommandEvents_Event) == t)
            {
                EventInfo ev = typeof(ICommandEvents_Event).GetEvent(eventName);
                return ev;
            }
            if (typeof (IEventEvents_Event) == t)
            {
                EventInfo ev = typeof (IEventEvents_Event).GetEvent(eventName);
                return ev;
            }
            return null;
        }
        public static EventInfo getEventInfo(string _type, string eventName)
        {

            if (_type == "ICommandEvents_Event")
            {
                return getEventInfo(typeof (ICommandEvents_Event), eventName);
            }
            if (_type == "IEventEvents_Event")
            {
                return getEventInfo(typeof (IEventEvents_Event), eventName);
            }
            return null;
        }
    }
}
