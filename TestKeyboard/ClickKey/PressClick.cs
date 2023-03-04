using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace TestKeyboard.ClickKey
{
    public class PressClick
    {
        [DllImport("user32.dll")]
        public static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData,  int dwExtraInfo);
        const int MOUSEEVENTF_MOVE = 0x0001; //移动鼠标
        const int MOUSEEVENTF_LEFTDOWN = 0x0002; //模拟鼠标左键按下
        const int MOUSEEVENTF_LEFTUP = 0x0004; //模拟鼠标左键抬起
        const int MOUSEEVENTF_RIGHTDOWN = 0x0008; //模拟鼠标右键按下
        const int MOUSEEVENTF_RIGHTUP = 0x0010; //模拟鼠标右键抬起
        const int MOUSEEVENTF_MIDDLEDOWN = 0x0020; //模拟鼠标中键按下
        const int MOUSEEVENTF_MIDDLEUP = 0x0040; //模拟鼠标中键抬起
        const int MOUSEEVENTF_ABSOLUTE = 0x8000; //标示是否采用绝对坐标

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetCursorPos(out Point point);
        [DllImport("user32.dll")]
        public static extern int GetSystemMetrics(SystemMetric smIndex);
        public void Click(int x,int y, EnumMouse clickType)
        {
            var clientX = GetSystemMetrics(SystemMetric.SM_CXSCREEN);
            var clientY = GetSystemMetrics(SystemMetric.SM_CYSCREEN); 
            if(clickType == EnumMouse.鼠标左键)
            {
                mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE, x * 65535 / clientX, y * 65535 / clientY, 0, 0);
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, x * 65535 / clientX, y * 65535 / clientY, 0, 0);
            }
            else if (clickType == EnumMouse.鼠标右键)
            {
                mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE, x * 65535 / clientX, y * 65535 / clientY, 0, 0);
                mouse_event(MOUSEEVENTF_RIGHTDOWN | MOUSEEVENTF_RIGHTUP, x * 65535 / clientX, y * 65535 / clientY, 0, 0);
            }
            else if (clickType == EnumMouse.鼠标右键)
            {
                mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE, x * 65535 / clientX, y * 65535 / clientY, 0, 0);
                mouse_event(MOUSEEVENTF_MIDDLEDOWN | MOUSEEVENTF_MIDDLEUP, x * 65535 / clientX, y * 65535 / clientY, 0, 0);
            }
           
        }
        public Point GetClickPos()
        {
            Point point;
            GetCursorPos(out point);
            return point;
        }
    }
}
