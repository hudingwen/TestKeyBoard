using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using TestKeyboard.ClickKey;
using TestKeyboard.SetForegroundWindow;

namespace TestKeyboard.Screen
{
    public class ScreenCat
    {
        //这里是调用 Windows API函数来进行截图
        //首先导入库文件
        [DllImportAttribute("gdi32.dll")]

        //声明函数
        private static extern IntPtr CreateDC
        (
            string Driver,   //驱动名称
            string Device,   //设备名称
            string Output,   //无用，可以设定为null
            IntPtr PrintData //任意的打印机数据
         );


        [DllImportAttribute("gdi32.dll")]
        private static extern bool BitBlt(
            IntPtr hdcDest,     //目标设备的句柄
            int XDest,          //目标对象的左上角X坐标
            int YDest,          //目标对象的左上角的Y坐标
            int Width,          //目标对象的宽度
            int Height,         //目标对象的高度
            IntPtr hdcScr,      //源设备的句柄
            int XScr,           //源设备的左上角X坐标
            int YScr,           //源设备的左上角Y坐标
            Int32 drRop         //光栅的操作值

            );
        /// <summary>
        /// 截图
        /// </summary>
        /// <param name="isFullScreen">是否全屏</param>
        /// <param name="posX">起始X坐标</param>
        /// <param name="posY">起始Y坐标</param>
        /// <param name="cWidth">截图宽度</param>
        /// <param name="cHeight">截图高度</param>
        /// <returns></returns>
        public static Image GetScreen(bool isFullScreen=true,int posX=0,int posY = 0, int cWidth=200,int cHeight = 100)
        {
            //创建显示器的DC
            IntPtr dcScreen = CreateDC("DISPLAY", null, null, (IntPtr)null);

            //由一个指定设备的句柄创建一个新的Graphics对象
            Graphics g1 = Graphics.FromHdc(dcScreen);
            Image MyImage = null;
            int tmpWidth, tmpHeigth;


            //获得保存图片的质量
            long level = 50; //0-100


            //如果是全屏捕获 
            if (isFullScreen)
            {
                //tmpWidth = 0;//屏幕宽度
                //tmpHeigth = 0;//屏幕高度
                tmpWidth = PressClick.GetSystemMetrics(SystemMetric.SM_CXSCREEN);
                tmpHeigth = PressClick. GetSystemMetrics(SystemMetric.SM_CYSCREEN);
                MyImage = new Bitmap(tmpWidth, tmpHeigth, g1);

                //创建位图图形对象
                Graphics g2 = Graphics.FromImage(MyImage);
                //获得窗体的上下文设备
                IntPtr dc1 = g1.GetHdc();
                //获得位图文件的上下文设备
                IntPtr dc2 = g2.GetHdc();

                //写入到位图
                BitBlt(dc2, 0, 0, tmpWidth, tmpHeigth, dc1, 0, 0, 13369376);

                //释放窗体的上下文设备
                g1.ReleaseHdc(dc1);
                //释放位图的上下文设备
                g2.ReleaseHdc(dc2);


                //保存图像并显示
                //SaveImageWithQuality(MyImage, level);
                //this.pictureBox1.Image = MyImage;

            }
            else
            {
                int X = Convert.ToInt32(posX);//起始X坐标
                int Y = Convert.ToInt32(posY);//起始Y坐标
                int Width = Convert.ToInt32(cWidth);//截图宽度
                int Height = Convert.ToInt32(cHeight);//截图高度

                MyImage = new Bitmap(Width, Height, g1);
                Graphics g2 = Graphics.FromImage(MyImage);
                IntPtr dc1 = g1.GetHdc();
                IntPtr dc2 = g2.GetHdc();

                BitBlt(dc2, 0, 0, Width, Height, dc1, X, Y, 13369376);
                g1.ReleaseHdc(dc1);
                g2.ReleaseHdc(dc2);

                //SaveImageWithQuality(Myimage, level);
                //this.pictureBox1.Image = Myimage; 
            }
            return MyImage;
        }
        /// <summary>
        /// 保存图片
        /// </summary>
        /// <param name="savePath">保存路径</param>
        /// <param name="bmp">图片</param>
        /// <param name="level">图片质量(0-100)</param>
        public static void SaveImageWithQuality(string savePath,Image bmp, long level)
        { 
            ImageCodecInfo jgpEncoder = GetEncoder(ImageFormat.Png);
            System.Drawing.Imaging.Encoder myEncoder =
                System.Drawing.Imaging.Encoder.Quality;
            EncoderParameters myEncoderParameters = new EncoderParameters(1);
            EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, level);
            myEncoderParameters.Param[0] = myEncoderParameter;
            bmp.Save(savePath, jgpEncoder, myEncoderParameters);



        } 

        private static ImageCodecInfo GetEncoder(ImageFormat format)
        {

            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();

            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }

        //窗口移动相关
        [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)]
        internal static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall, ExactSpelling = true, SetLastError = true)]
        internal static extern bool GetWindowRect(IntPtr hWnd, ref RECT rect);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall, ExactSpelling = true, SetLastError = true)]
        internal static extern void MoveWindow(IntPtr hwnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);


        public static bool Move(string windowName,int x=0,int y=0)
        {
            var mFocus = new Foreground();
            IntPtr id;
            RECT Rect = new RECT();
            Thread.Sleep(2000);
            id = mFocus.FindWindow(windowName);
            if (id == IntPtr.Zero) return false;
            GetWindowRect(id, ref Rect);
            MoveWindow(id, x, y, Rect.right - Rect.left, Rect.bottom - Rect.top, true);
            return true;
        }
    }
    internal struct RECT
    {
        public int left;
        public int top;
        public int right;
        public int bottom;
    }

   
}
