using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace TestKeyboard.SetForegroundWindow
{
   
    public class Foreground
    {
        private delegate bool WNDENUMPROC(IntPtr hWnd, int lParam);
        WindowInfo[] infos = { };
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int x, int y, int Width, int Height, int flags);

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr ProcessId);
        [DllImport("user32.dll")]
        static extern bool AttachThreadInput(uint idAttach, uint idAttachTo,bool fAttach);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(WNDENUMPROC lpEnumFunc, int lParam);

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll", EntryPoint = "ShowWindow", CharSet = CharSet.Auto)]
        public static extern int ShowWindow(IntPtr hwnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern IntPtr AttachThreadInput(IntPtr idAttach, IntPtr idAttachTo, int fAttach);
        [DllImport("user32.dll")]
        static extern bool AllowSetForegroundWindow(uint dwProcessId);
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("kernel32.dll")]
        static extern uint GetCurrentThreadId();
        public WindowInfo[] GetAllDesktopWindows()
        {

            //用来保存窗口对象 列表
            List<WindowInfo> wndList = new List<WindowInfo>();

            //enum all desktop windows 
            EnumWindows(delegate (IntPtr hWnd, int lParam)
            {
                WindowInfo wnd = new WindowInfo();
                StringBuilder sb = new StringBuilder(256);

                //get hwnd 
                wnd.hWnd = hWnd;

                //get window name  
                GetWindowText(hWnd, sb, sb.Capacity);

                wnd.szWindowName = sb.ToString();

                //wnd.processID = GetWindowThreadProcessId(hWnd, hWnd);
                //add it into list 

                wndList.Add(wnd);
                return true;
            }, 0);

            return wndList.ToArray();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="windowName"></param>
        /// <param name="nCmdShow">1-恢复正常窗口大小 2-最小化窗口 3-最大化窗口</param>
        /// <returns></returns>
        public bool ShowWindow(string windowName,int nCmdShow)
        {
            
            IntPtr hWnd;
            int result;

            hWnd = IntPtr.Zero;
            infos = GetAllDesktopWindows();
            foreach (WindowInfo info in infos)
                if (info.szWindowName == windowName)
                    hWnd = info.hWnd;

            if (hWnd == IntPtr.Zero)
            {

                return false;
            }

            try
            {
                result = ShowWindow(hWnd, nCmdShow);

                if (result == 1)
                    return true;
                else
                    return false;
            }

            catch
            {
                return false;
            }
        }
        int HWND_TOP = 0;
        int HWND_BOTTOM = 1;
        int HWND_TOPMOST = -1;
        int HWND_NOTOPMOST = -2; 



       int SWP_NOSIZE = 1; 
       int SWP_NOMOVE = 2; 
       int SWP_NOZORDER = 4; 
       int SWP_NOREDRAW = 8;
       int SWP_NOACTIVATE = 0x10; 
       int SWP_FRAMECHANGED = 0x20; 
       int SWP_SHOWWINDOW = 0x40; 
       int SWP_HIDEWINDOW = 0x80; 
       int SWP_NOCOPYBITS = 0x100;
       int SWP_NOOWNERZORDER = 0x200;
       int SWP_NOSENDCHANGING = 0x400;


        int SW_HIDE = 0;
        int SW_NORMAL = 1;
        int SW_SHOWNORMAL = 1;
        int SW_SHOWMINIMIZED = 2;
        int SW_SHOWMAXIMIZED = 3;
        int SW_MAXIMIZE = 3;
        int SW_SHOWNOACTIVATE = 4;
        int SW_SHOW = 5;
        int SW_MINIMIZE = 6;
        int SW_SHOWMINNOACTIVE = 7;
        int SW_SHOWNA = 8;
        int SW_RESTORE = 9;
        int SW_SHOWDEFAULT = 10;
        int SW_FORCEMINIMIZE = 11;
        public bool FocusWindow(string windowName,ref string msg)
        {
            IntPtr hWnd;
            int result;

            hWnd = IntPtr.Zero;
            infos = GetAllDesktopWindows();
            WindowInfo findWindow;
            foreach (WindowInfo info in infos)
            {
                if (info.szWindowName == windowName)
                {
                    hWnd = info.hWnd;
                    findWindow = info;
                    break;
                }
            }

            if (hWnd == IntPtr.Zero)
            {
                msg = "未找到程序";
                return false;
            }

            try
            {
                


                var hForeWnd = GetForegroundWindow();
                var dwForeID = GetWindowThreadProcessId(hForeWnd, IntPtr.Zero);
                var dwCurID = GetCurrentThreadId();
                AttachThreadInput(dwCurID, dwForeID, true);
                ShowWindow(hWnd, SW_SHOWNORMAL);
                SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE);
                SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE);
                SetForegroundWindow(hWnd);
                AttachThreadInput(dwCurID, dwForeID, false);

                msg = "找到窗口";
                var isok = true; 
                return true;


                 
                Thread.Sleep(200);
                result = ShowWindow(hWnd,1);
                Thread.Sleep(200);
                if (result <= 0) {
                    msg = "失败-1";
                    return false;
                }
                Thread.Sleep(200);
                result = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE);
                if (result <= 0)
                {
                    msg = "失败-2";
                    return false;
                }
                Thread.Sleep(200);
                result = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE);
                if (result <= 0)
                {
                    msg = "失败-3";
                    return false;
                }
                Thread.Sleep(200);
                isok =  SetForegroundWindow(hWnd);
               
                if (!isok)
                {
                    msg = "失败-4";
                    return false;
                }

                //isok = AttachThreadInput((uint)AppDomain.GetCurrentThreadId(), curInfo.processID, false);
               // msg = "5";
                //if (!isok) return false;

                return isok;


                //result = 1;

                //SetForegroundWindow(hWnd);

                //显示
                //result = SetWindowPos(hWnd, 0, 0, 0, 0, 0, 0x001 | 0x002 | 0x004 | 0x040);
                //置顶
                //result = SetWindowPos(hWnd, -1, 0, 0, 0, 0, 0x001 | 0x002 | 0x040);

                if (result == 1)
                    return true;
                else
                    return false;
            }

            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 找到窗口与信息
        /// </summary>
        /// <param name="windowName"></param>
        /// <param name="msg"></param>
        /// <returns></returns>
        public IntPtr FindWindow(string windowName)
        {
            infos = GetAllDesktopWindows();
            foreach (WindowInfo info in infos)
            {
                if (info.szWindowName == windowName)
                {
                    return info.hWnd;
                }
            }
            return IntPtr.Zero;
        }
    }
}
