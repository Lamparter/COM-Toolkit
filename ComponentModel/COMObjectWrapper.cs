using System;
using System.Collections.Generic;
using System.Text;

namespace Riverside.ComponentModel
{
    public class COMObjectWrapper<T> : IDisposable where T : class
    {
        private T _comObject;

        public COMObjectWrapper(string progId)
        {
            _comObject = InteropUtilities.CreateCOMObject<T>(progId);
        }

        public T COMObject => _comObject;

        public void Dispose()
        {
            InteropUtilities.ReleaseCOMObject(_comObject);
            _comObject = null;
        }
    }
}
