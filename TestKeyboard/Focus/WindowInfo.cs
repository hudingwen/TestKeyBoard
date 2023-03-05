using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestKeyboard.SetForegroundWindow
{
    public struct WindowInfo

    {
        public IntPtr hWnd { get; set; }
        public string szWindowName{ get; set; }
    }
}
