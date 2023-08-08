using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace AutoUpDataBoss
{
    class Mouse
    {
        public static void denglu()
        {
            Mouse.MouseMoveToPoint(330, 400);
            Mouse.WaitFunctions(1);
            Mouse.LeftClick();
            Mouse.WaitFunctions(1);
            Mouse.MouseMoveToPoint(300, 450);
            Mouse.WaitFunctions(1);
            Mouse.LeftClick();
            Mouse.WaitFunctions(1);
            Mouse.MouseMoveToPoint(100, 215);
            Mouse.WaitFunctions(1);
            Mouse.LeftClick();
            Mouse.WaitFunctions(1);
            Mouse.KeyboardInput("jy_szw");
            Mouse.WaitFunctions(1);
            Mouse.MouseMoveToPoint(100, 315);
            Mouse.WaitFunctions(1);
            Mouse.LeftClick();
            Mouse.WaitFunctions(1);
            Mouse.KeyboardInput("Aa@6543212");
            Mouse.WaitFunctions(1);
            Mouse.MouseMoveToPoint(100, 380);
            Mouse.WaitFunctions(1);
            Mouse.LeftClick();
            Mouse.WaitFunctions(5);
        }


        //调用鼠标事件WindoswsAPI,具体请参见 http://www.office-cn.net/t/api/index.html?mouse_event.htm
        [DllImport("user32")]
        private static extern int mouse_event(int flag, int dx, int dy, int delta, int info);

        public enum MouseFlags
        {
            //移动鼠标 
            MouseMove = 0x0001,
            //模拟鼠标左键按下 
            MouseLeftDown = 0x0002,
            //模拟鼠标左键抬起 
            MouseLeftUp = 0x0004,
            //模拟鼠标右键按下 
            MouseRightDown = 0x0008,
            //模拟鼠标右键抬起 
            MouseRightUp = 0x0010,
            //模拟鼠标中键按下 
            MouseMiddleDown = 0x0020,
            //模拟鼠标中键抬起 
            MouseMiddleUp = 0x0040,
            //标示是否采用绝对坐标 
            IsAbsolute = 0x8000,
            //模拟鼠标滚轮滚动操作，滚轮滑动数值为形参delta的值
            MouseWheel = 0x0800,
        }

        //调用鼠标光标位置WindoswsAPI,具体请参见 http://www.office-cn.net/t/api/index.html?setcursorpos.htm
        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int x, int y);


        /// <summary>
        /// 移动光标到指定位置
        /// </summary>
        /// <param name="x">屏幕水平方向X坐标</param>
        /// <param name="y">屏幕垂直方向Y坐标,注意对于Y坐标屏幕原点(0,0)为左上角</param>
        public static bool MouseMoveToPoint(int x, int y)
        {
            return SetCursorPos(x, y);
            //当然也可以使用Cursor.Position = new System.Drawing.Point(x, y);
        }

        /// <summary>
        /// 鼠标左按下
        /// </summary>
        public static void LeftDown()
        {
            mouse_event((int)MouseFlags.MouseLeftDown, 0, 0, 0, 0);
        }
        /// <summary>
        /// 鼠标左抬起
        /// </summary>
        public static void LeftUp()
        {
            mouse_event((int)MouseFlags.MouseLeftUp, 0, 0, 0, 0);
        }
        /// <summary>
        /// 鼠标左击
        /// </summary>
        public static void LeftClick()
        {
            LeftDown();
            LeftUp();
        }

        /// <summary>
        /// 鼠标中键按下
        /// </summary>
        public static void MiddleDown()
        {
            mouse_event((int)MouseFlags.MouseMiddleDown, 0, 0, 0, 0);
        }
        /// <summary>
        /// 鼠标中键抬起
        /// </summary>
        public static void MiddleUp()
        {
            mouse_event((int)MouseFlags.MouseMiddleUp, 0, 0, 0, 0);
        }

        /// <summary>
        /// 滚轮滚轮点击
        /// </summary>
        public static void MiddleClick()
        {
            MiddleDown();
            MiddleUp();
        }

        /// <summary>
        /// 鼠标滚轮滑动
        /// </summary>
        /// <param name="delta">滑动数值</param>
        public static void MiddleWheel(int delta)
        {
            mouse_event((int)MouseFlags.MouseWheel, 0, 0, delta, 0);
        }

        /// <summary>
        /// 鼠标右键按下
        /// </summary>
        public static void RightDown()
        {
            mouse_event((int)MouseFlags.MouseRightDown, 0, 0, 0, 0);
        }
        /// <summary>
        /// 鼠标中键抬起
        /// </summary>
        public static void RightUp()
        {
            mouse_event((int)MouseFlags.MouseRightUp, 0, 0, 0, 0);
        }

        /// <summary>
        /// 滚轮滚轮点击
        /// </summary>
        public static void RightClick()
        {
            RightDown();
            RightUp();
        }

        /// <summary>
        /// 键盘输入字符串
        /// </summary>
        /// <param name="value"></param>
        /// <param name="is_enter">默认不开启回车键</param>
        public static void KeyboardInput(string value/*, bool is_enter = false*/)
        {
            //if (is_enter)
            //    value += "{ENTER}";
            //注意要使得SendKeys起作用需要在App.config文件中configuration一栏配置使用信息
            //具体请使用请参见 https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.sendkeys.sendwait?view=netframework-4.5
            SendKeys.SendWait(value);
        }


        public static void WaitFunctions(int waitTime)
        {
            if (waitTime <= 0) return;
            DateTime nowTimer = DateTime.Now;
            int interval = 0;
            while (interval < waitTime)
            {
                TimeSpan spand = DateTime.Now - nowTimer;
                interval = spand.Seconds;
            }
        }

    }
}
