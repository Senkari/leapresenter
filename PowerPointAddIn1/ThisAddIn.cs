using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Runtime.InteropServices;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn1
{ 
    public partial class ThisAddIn
    {
        //private:
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(IntPtr hWnd, ref RECT lpRect);
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private Overlay overlayWindow;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.SlideShowBegin += new PowerPoint.EApplication_SlideShowBeginEventHandler(Application_SlideShowStarted);
            this.Application.SlideShowEnd += new PowerPoint.EApplication_SlideShowEndEventHandler(Application_SlideShowEnded);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        private void Application_SlideShowStarted(PowerPoint.SlideShowWindow window)
        {
            overlayWindow = new Overlay();
            overlayWindow.setSlideShowWindow(window);
            overlayWindow.setSlideShowActive(true);
            RECT rect = new RECT();
            GetWindowRect(new IntPtr(window.HWND), ref rect);
            overlayWindow.Left = rect.Left; // EDIT: do not use slideshowWindow.Left, etc.
            overlayWindow.Top = rect.Top;
            overlayWindow.Width = rect.Right;
            overlayWindow.Height = rect.Bottom;
            overlayWindow.Show();
            //Wn.View.PointerType = PowerPoint.PpSlideShowPointerType.ppSlideShowPointerPen;
            //Wn.View.DrawLine(0, 0, 20, 20);
        }

        private void Application_SlideShowEnded(PowerPoint.Presentation presentation)
        {
            overlayWindow.setSlideShowActive(false);
            overlayWindow.Close();
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
