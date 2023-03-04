using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestKeyboard.MonitorEvent
{
    /// <summary>
    /// 钩子委托声明
    /// </summary>
    /// <param name="nCode"></param>
    /// <param name="wParam"></param>
    /// <param name="lParam"></param>
    /// <returns></returns>
    public delegate int HookProc(int nCode, Int32 wParam, IntPtr lParam);

    /// <summary>
    /// 鼠标更新事件委托声明
    /// </summary>
    /// <param name="x">x坐标</param>
    /// <param name="y">y坐标</param>
    public delegate void MouseUpdateEventHandler(int x, int y);

    /// <summary>
    /// 无返回委托声明
    /// </summary>
    public delegate void VoidCallback();
    class KeyMouseHook
    {
        /// <summary>
        /// 鼠标点击事件
        /// </summary>
        public event MouseEventHandler OnMouseActivity;

        /// <summary>
        /// 鼠标更新事件
        /// </summary>
        /// <remarks>当鼠标移动或者滚轮滚动时触发</remarks>
        public event MouseUpdateEventHandler OnMouseUpdate;

        /// <summary>
        /// 按键按下事件
        /// </summary>
        public event KeyEventHandler OnKeyDown;

        /// <summary>
        /// 按键按下并释放事件
        /// </summary>
        public event KeyPressEventHandler OnKeyPress;

        /// <summary>
        /// 按键释放事件
        /// </summary>
        public event KeyEventHandler OnKeyUp;




        public delegate int HookProc(int nCode, Int32 wParam, IntPtr lParam);
        static int hKeyboardHook = 0; //声明键盘钩子处理的初始值
        //值在Microsoft SDK的Winuser.h里查询
        public const int WH_MOUSE_LL = (int)HookType.WH_MOUSE_LL;   //线程键盘钩子监听鼠标消息设为2，全局键盘监听鼠标消息设为13
        HookProc KeyboardHookProcedure; //声明KeyboardHookProcedure作为HookProc类型
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        /// <summary>
        /// 鼠标钩子事件结构定义
        /// </summary>
        /// <remarks>详细说明请参考MSDN中关于 MSLLHOOKSTRUCT 的说明</remarks>
        [StructLayout(LayoutKind.Sequential)]
        public struct MouseHookStruct
        {
            /// <summary>
            /// Specifies a POINT structure that contains the x- and y-coordinates of the cursor, in screen coordinates.
            /// </summary>
            public POINT Point;

            public UInt32 MouseData;
            public UInt32 Flags;
            public UInt32 Time;
            public UInt32 ExtraInfo;
        }
        //使用此功能，安装了一个钩子
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern int SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hInstance, int threadId);


        //调用此函数卸载钩子
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern bool UnhookWindowsHookEx(int idHook);


        //使用此功能，通过信息钩子继续下一个钩子
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern int CallNextHookEx(int idHook, int nCode, Int32 wParam, IntPtr lParam);

        // 取得当前线程编号（线程钩子需要用到）
        [DllImport("kernel32.dll")]
        static extern int GetCurrentThreadId();

        //使用WINDOWS API函数代替获取当前实例的函数,防止钩子失效
        [DllImport("kernel32.dll")]
        public static extern IntPtr GetModuleHandle(string name);

        public void Start()
        {
            // 安装键盘钩子
            if (hKeyboardHook == 0)
            {
                KeyboardHookProcedure = new HookProc(KeyboardHookProc);
                //hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, KeyboardHookProcedure, GetModuleHandle(System.Diagnostics.Process.GetCurrentProcess().MainModule.ModuleName), 0);
                //hKeyboardHook = SetWindowsHookEx(WH_MOUSE_LL, KeyboardHookProcedure, Marshal.GetHINSTANCE(Assembly.GetExecutingAssembly().GetModules()[0]), 0);
                //************************************
                //键盘线程钩子
                //SetWindowsHookEx(13, KeyboardHookProcedure, IntPtr.Zero, GetCurrentThreadId());//指定要监听的线程idGetCurrentThreadId(),
                //键盘全局钩子,需要引用空间(using System.Reflection;)
                hKeyboardHook = SetWindowsHookEx(WH_MOUSE_LL, KeyboardHookProcedure, Marshal.GetHINSTANCE(Assembly.GetExecutingAssembly().GetModules()[0]),0);
                //
                //关于SetWindowsHookEx (int idHook, HookProc lpfn, IntPtr hInstance, int threadId)函数将钩子加入到钩子链表中，说明一下四个参数：
                //idHook 钩子类型，即确定钩子监听何种消息，上面的代码中设为2，即监听键盘消息并且是线程钩子，如果是全局钩子监听键盘消息应设为13，
                //线程钩子监听鼠标消息设为7，全局钩子监听鼠标消息设为14。lpfn 钩子子程的地址指针。如果dwThreadId参数为0 或是一个由别的进程创建的
                //线程的标识，lpfn必须指向DLL中的钩子子程。 除此以外，lpfn可以指向当前进程的一段钩子子程代码。钩子函数的入口地址，当钩子钩到任何
                //消息后便调用这个函数。hInstance应用程序实例的句柄。标识包含lpfn所指的子程的DLL。如果threadId 标识当前进程创建的一个线程，而且子
                //程代码位于当前进程，hInstance必须为NULL。可以很简单的设定其为本应用程序的实例句柄。threaded 与安装的钩子子程相关联的线程的标识符
                //如果为0，钩子子程与所有的线程关联，即为全局钩子
                //************************************
                //如果SetWindowsHookEx失败
                if (hKeyboardHook == 0)
                {
                    Stop();
                    throw new Exception("安装键盘钩子失败");
                }
            }
        }
        public void Stop()
        {
            bool retKeyboard = true;


            if (hKeyboardHook != 0)
            {
                retKeyboard = UnhookWindowsHookEx(hKeyboardHook);
                hKeyboardHook = 0;
            }

            if (!(retKeyboard)) throw new Exception("卸载钩子失败！");
        }
        //ToAscii职能的转换指定的虚拟键码和键盘状态的相应字符或字符
        [DllImport("user32")]
        public static extern int ToAscii(int uVirtKey, //[in] 指定虚拟关键代码进行翻译。
                                         int uScanCode, // [in] 指定的硬件扫描码的关键须翻译成英文。高阶位的这个值设定的关键，如果是（不压）
                                         byte[] lpbKeyState, // [in] 指针，以256字节数组，包含当前键盘的状态。每个元素（字节）的数组包含状态的一个关键。如果高阶位的字节是一套，关键是下跌（按下）。在低比特，如果设置表明，关键是对切换。在此功能，只有肘位的CAPS LOCK键是相关的。在切换状态的NUM个锁和滚动锁定键被忽略。
                                         byte[] lpwTransKey, // [out] 指针的缓冲区收到翻译字符或字符。
                                         int fuState); // [in] Specifies whether a menu is active. This parameter must be 1 if a menu is active, or 0 otherwise.

        //获取按键的状态
        [DllImport("user32")]
        public static extern int GetKeyboardState(byte[] pbKeyState);


        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        private static extern short GetKeyState(int vKey);

        /// <summary>
        /// 鼠标开始
        /// </summary>
        private const int WM_MOUSEFIRST = 0x200;
        /// <summary>
        /// 鼠标移动
        /// </summary>
        private const int WM_MOUSEMOVE = 0x200;
        /// <summary>
        /// 左键按下
        /// </summary>
        private const int WM_LBUTTONDOWN = 0x201;
        /// <summary>
        /// 左键释放
        /// </summary>
        private const int WM_LBUTTONUP = 0x202;
        /// <summary>
        /// 左键双击
        /// </summary>
        private const int WM_LBUTTONDBLCLK = 0x203;
        /// <summary>
        /// 右键按下
        /// </summary>
        private const int WM_RBUTTONDOWN = 0x204;
        /// <summary>
        /// 右键释放
        /// </summary>
        private const int WM_RBUTTONUP = 0x205;
        /// <summary>
        /// 右键双击
        /// </summary>
        private const int WM_RBUTTONDBLCLK = 0x206;
        /// <summary>
        /// 中键按下
        /// </summary>
        private const int WM_MBUTTONDOWN = 0x207;
        /// <summary>
        /// 中键释放
        /// </summary>
        private const int WM_MBUTTONUP = 0x208;
        /// <summary>
        /// 中键双击
        /// </summary>
        private const int WM_MBUTTONDBLCLK = 0x209;
        /// <summary>
        /// 滚轮滚动
        /// </summary>
        private const int WM_MOUSEWHEEL = 0x020A;

        private int KeyboardHookProc(int nCode, Int32 wParam, IntPtr lParam)
        {
            //鼠标移动事件
            if ((nCode >= 0) && (this.OnMouseUpdate != null) && (wParam == (int)WM_MOUSEMOVE || wParam == (int)WM_MOUSEWHEEL))
            {
                MouseHookStruct MouseInfo = (MouseHookStruct)Marshal.PtrToStructure(lParam, typeof(MouseHookStruct));
                this.OnMouseUpdate(MouseInfo.Point.X, MouseInfo.Point.Y);
            }

            //鼠标点击事件
            if ((nCode >= 0) && (this.OnMouseActivity != null) && (wParam == (int)WM_RBUTTONDOWN || wParam == (int)WM_MBUTTONDOWN || wParam == (int)WM_LBUTTONDOWN))
            {
                MouseButtons button = MouseButtons.None;
                int clickCount = 0;

                switch (wParam)
                {
                    case (int)WM_MBUTTONDOWN:
                        button = MouseButtons.Middle;
                        clickCount = 1;
                        break;
                    case (int)WM_MBUTTONUP:
                        button = MouseButtons.Middle;
                        clickCount = 0;
                        break;
                }

                MouseHookStruct MouseInfo = (MouseHookStruct)Marshal.PtrToStructure(lParam, typeof(MouseHookStruct));
                MouseEventArgs mouseEvent = new MouseEventArgs(button, clickCount, MouseInfo.Point.X, MouseInfo.Point.Y, 0);
                this.OnMouseActivity(this, mouseEvent);
            }

            ////如果正常运行并且用户要监听鼠标的消息  
            //if ((nCode >= 0) && (OnMouseActivity != null))
            //{
            //    MouseButtons button = MouseButtons.None;
            //    int clickCount = 0;
            //
            //    switch (wParam)
            //    {
            //        case WM_MOUSE.WM_MBUTTONDOWN:
            //            button = MouseButtons.Middle;
            //            clickCount = 1;
            //            break;
            //        case WM_MOUSE.WM_MBUTTONUP:
            //            button = MouseButtons.Middle;
            //            clickCount = 0;
            //            break;
            //    }
            //
            //    //从回调函数中得到鼠标的信息  
            //    MouseHookStruct MyMouseHookStruct = (MouseHookStruct)Marshal.PtrToStructure(lParam, typeof(MouseHookStruct));
            //    MouseEventArgs e = new MouseEventArgs(button, clickCount, MyMouseHookStruct.pt.x, MyMouseHookStruct.pt.y, 0);
            //    //if(e.X>700)return 1;//如果想要限制鼠标在屏幕中的移动区域可以在此处设置  
            //    OnMouseActivity(this, e);
            //}
            //如果返回1，则结束消息，这个消息到此为止，不再传递。
            //如果返回0或调用CallNextHookEx函数则消息出了这个钩子继续往下传递，也就是传给消息真正的接受者
            return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
        }
        ~KeyMouseHook()
        {
            Stop();
        }
    }
}
