using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace Riverside.ComponentModel
{
    public static class COMErrorHandling
    {
        public static void CheckHResult(int hResult)
        {
            if (hResult < 0)
            {
                Marshal.ThrowExceptionForHR(hResult);
            }
        }
    }
}
