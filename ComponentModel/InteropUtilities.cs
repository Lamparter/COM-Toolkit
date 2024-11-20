using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace Riverside.ComponentModel
{
    public static class InteropUtilities
    {
        private static readonly Dictionary<Delegate, int> EventHandlerCookies = new Dictionary<Delegate, int>();

        public static void ReleaseCOMObject(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }

        public static T CreateCOMObject<T>(string progId) where T : class
        {
            Type comType = Type.GetTypeFromProgID(progId);
            return comType == null ? throw new ArgumentException($"ProgID {progId} not found.") : Activator.CreateInstance(comType) as T;
        }

        public static void AttachEventHandler<T>(object comObject, string eventName, Delegate handler)
        {
            if (comObject == null)
            {
                throw new ArgumentNullException(nameof(comObject));
            }

            if (string.IsNullOrEmpty(eventName))
            {
                throw new ArgumentException("Event name cannot be null or empty.", nameof(eventName));
            }

            if (handler == null)
            {
                throw new ArgumentNullException(nameof(handler));
            }

            if (!(comObject is IConnectionPointContainer connectionPointContainer))
            {
                throw new ArgumentException("COM object does not support event handling.", nameof(comObject));
            }

            Guid guid = typeof(T).GUID;
            connectionPointContainer.FindConnectionPoint(ref guid, out IConnectionPoint connectionPoint);
            connectionPoint.Advise(handler, out int cookie);
            EventHandlerCookies[handler] = cookie;
        }

        public static void DetachEventHandler<T>(object comObject, string eventName, Delegate handler)
        {
            if (comObject == null)
            {
                throw new ArgumentNullException(nameof(comObject));
            }

            if (string.IsNullOrEmpty(eventName))
            {
                throw new ArgumentException("Event name cannot be null or empty.", nameof(eventName));
            }

            if (handler == null)
            {
                throw new ArgumentNullException(nameof(handler));
            }

            IConnectionPointContainer connectionPointContainer = comObject as IConnectionPointContainer;
            if (connectionPointContainer == null)
            {
                throw new ArgumentException("COM object does not support event handling.", nameof(comObject));
            }

            if (!EventHandlerCookies.TryGetValue(handler, out int cookie))
            {
                throw new ArgumentException("Event handler not found.", nameof(handler));
            }

            Guid guid = typeof(T).GUID;
            connectionPointContainer.FindConnectionPoint(ref guid, out IConnectionPoint connectionPoint);
            connectionPoint.Unadvise(cookie);
            EventHandlerCookies.Remove(handler);
        }
    }
}
