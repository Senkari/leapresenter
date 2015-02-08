using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Runtime.InteropServices;
using System.Windows;

namespace PowerPointAddIn1
{
    class MouseCursor
    {
       
        [DllImport("user32.dll")]

        private static extern bool SetCursorPos(int x, int y);
        
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]

        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, UIntPtr dwExtraInfo);

        public static void setCursor(int x, int y)
        {
            SetCursorPos(x, y);
        }

        public static void sendLeftMouseDown()
        {
            mouse_event(0x02, 0, 0, 0, UIntPtr.Zero);
        }

        public static void sendLeftMouseUp()
        {
            mouse_event(0x04, 0, 0, 0, UIntPtr.Zero);
        }
    }
}
