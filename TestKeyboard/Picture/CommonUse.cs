using Emgu.CV.Structure;
using Emgu.CV.UI;
using Emgu.CV.Util;
using Emgu.CV;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestKeyboard.Picture
{
    internal class CommonUse
    {
        /// <summary>
        /// 获取给定图像的最大矩形边界
        /// </summary>
        /// <param name="src"></param>
        /// <returns></returns>
        public VectorOfVectorOfPoint GetBoundaryOfPic(Mat src)
        {
            Mat dst = new Mat();
            Mat src_gray = new Mat();
            CvInvoke.CvtColor(src, src_gray, Emgu.CV.CvEnum.ColorConversion.Bgr2Gray);

            //边缘检测
            CvInvoke.Canny(src_gray, dst, 120, 180);

            //寻找答题卡矩形边界（最大的矩形）
            VectorOfVectorOfPoint contours = new VectorOfVectorOfPoint();//创建VectorOfVectorOfPoint数据类型用于存储轮廓

            CvInvoke.FindContours(dst, contours, null, Emgu.CV.CvEnum.RetrType.External,
                Emgu.CV.CvEnum.ChainApproxMethod.ChainApproxSimple);//提取轮廓

            VectorOfVectorOfPoint max_contour = new VectorOfVectorOfPoint();//用于存储筛选过后的轮廓

            int ksize = contours.Size; //获取连通区域个数
            if (ksize == 1)
            {
                max_contour = contours;
            }
            else
            {
                //double maxLength = -1;//用于保存轮廓周长的最大值
                double maxArea = -1;//面积
                int index = -1;//轮廓周长的最大值的序号
                for (int i = 0; i < ksize; i++)
                {
                    VectorOfPoint contour = contours[i];//获取独立的连通轮廓
                                                        //double length = CvInvoke.ArcLength(contour, true);//计算连通轮廓的周长

                    //if (length > maxLength)
                    //{
                    //    maxLength = length;
                    //    index = i;
                    //}

                    double area = CvInvoke.ContourArea(contour, false);
                    if (area > maxArea)
                    {
                        maxArea = area;
                        index = i;
                    }
                }
                max_contour.Push(contours[index]);//筛选后的连通轮廓
            }
            return max_contour;
        }

        /// <summary>
        /// 进行透视操作，获取矫正后图像
        /// </summary>
        /// <param name="src"></param>
        /// <param name="result_contour"></param>
        public Mat MyWarpPerspective(Mat src, VectorOfVectorOfPoint max_contour)
        {
            //拟合答题卡的几何轮廓,保存点集pts并顺时针排序
            VectorOfPoint pts = new VectorOfPoint();//用于存放逼近的结果
            VectorOfPoint tempContour = max_contour[0];//临时用
            double result_length = CvInvoke.ArcLength(tempContour, true);
            CvInvoke.ApproxPolyDP(tempContour, pts, result_length * 0.02, true); //几何逼近，获取矩形4个顶点坐标

            if (pts.Size != 4)
            {
                //最大轮廓不是矩形时，将原图转灰度图后返回
                Mat src_gray = new Mat();
                CvInvoke.CvtColor(src, src_gray, Emgu.CV.CvEnum.ColorConversion.Bgr2Gray);//转灰度图
                return src_gray;
            }
            else
            {
                //Point[]转换为PointF[]类型
                PointF[] pts_src = Array.ConvertAll(pts.ToArray(), new Converter<Point, PointF>(PointToPointF));

                //点集顺时针排序
                pts_src = SortPointsByClockwise(pts_src);

                //确定透视变换的宽度、高度
                Size sizeOfRect = CalSizeOfRect(pts_src);
                int width = sizeOfRect.Width;
                int height = sizeOfRect.Height;

                //计算透视变换矩阵
                PointF[] pts_target = new PointF[] { new PointF(0, 0), new PointF(width - 1, 0) ,
                        new PointF(width - 1, height - 1) ,new PointF(0, height - 1)};

                //计算透视矩阵
                Mat data = CvInvoke.GetPerspectiveTransform(pts_src, pts_target);
                //进行透视操作
                Mat src_gray = new Mat();
                Mat mat_Perspective = new Mat();
                CvInvoke.CvtColor(src, src_gray, Emgu.CV.CvEnum.ColorConversion.Bgr2Gray);
                CvInvoke.WarpPerspective(src_gray, mat_Perspective, data, new Size(width, height));

                return mat_Perspective;
            }
        }

        /// <summary>
        /// 将给定点集顺时针排序
        /// </summary>
        /// <param name="pts_src">四边形四个顶点组成的点集</param>
        /// <returns></returns>
        public PointF[] SortPointsByClockwise(PointF[] pts_src)
        {
            if (pts_src.Length != 4) return null;//确保为四边形

            //求四边形中心点？坐标
            float x_average = 0;
            float y_average = 0;
            float x_sum = 0;
            float y_sum = 0;
            for (int i = 0; i < 4; i++)
            {
                x_sum += pts_src[i].X;
                y_sum += pts_src[i].Y;
            }
            x_average = x_sum / 4;
            y_average = y_sum / 4;
            PointF center = new PointF(x_average, y_average);

            PointF[] result = new PointF[4];
            for (int i = 0; i < 4; i++)
            {
                if (pts_src[i].X < center.X && pts_src[i].Y < center.Y)
                {
                    result[0] = pts_src[i];//左上角点
                    continue;
                }
                if (pts_src[i].X > center.X && pts_src[i].Y < center.Y)
                {
                    result[1] = pts_src[i];//右上角点
                    continue;
                }
                if (pts_src[i].X > center.X && pts_src[i].Y > center.Y)
                {
                    result[2] = pts_src[i];//右下角点
                    continue;
                }
                if (pts_src[i].X < center.X && pts_src[i].Y > center.Y)
                {
                    result[3] = pts_src[i];//左下角点
                    continue;
                }
            }

            return result;
        }

        /// <summary>
        /// 计算给定四个坐标点四边形的宽、高
        /// </summary>
        /// <param name="pts_src"></param>
        /// <returns></returns>
        public Size CalSizeOfRect(PointF[] pts_src)
        {
            if (pts_src.Length != 4) return new Size(0, 0);//确保为四边形

            //点集顺时针排序
            pts_src = SortPointsByClockwise(pts_src);

            //确定透视变换的宽度、高度
            int width;
            int height;

            double width1 = Math.Pow(pts_src[0].X - pts_src[1].X, 2) + Math.Pow(pts_src[0].Y - pts_src[1].Y, 2);
            double width2 = Math.Pow(pts_src[2].X - pts_src[3].X, 2) + Math.Pow(pts_src[2].Y - pts_src[3].Y, 2);

            width = width1 > width2 ? (int)Math.Sqrt(width1) : (int)Math.Sqrt(width2);//根号下a方+b方，且取宽度最大的

            double height1 = Math.Pow(pts_src[0].X - pts_src[3].X, 2) + Math.Pow(pts_src[0].Y - pts_src[3].Y, 2);
            double height2 = Math.Pow(pts_src[1].X - pts_src[2].X, 2) + Math.Pow(pts_src[1].Y - pts_src[2].Y, 2);

            height = height1 > height2 ? (int)Math.Sqrt(height1) : (int)Math.Sqrt(height2);

            return new Size(width, height);
        }

        /// <summary>
        /// Point转换为PointF类型
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        public static PointF PointToPointF(Point p)
        {
            return new PointF(p.X, p.Y);
        }

        /// <summary>
        /// 形态学膨胀
        /// </summary>
        /// <param name="mat"></param>
        /// <returns></returns>
        public Mat MyDilate(Mat mat)
        {
            //1.膨胀，改善轮廓
            Mat struct_element = CvInvoke.GetStructuringElement(Emgu.CV.CvEnum.ElementShape.Cross,
                new Size(3, 3), new Point(-1, -1));//结构元素
            Mat mat_dilate = new Mat();
            CvInvoke.MorphologyEx(mat, mat_dilate, Emgu.CV.CvEnum.MorphOp.Dilate, struct_element, new Point(-1, -1), 1,
                Emgu.CV.CvEnum.BorderType.Default, new MCvScalar(0, 0, 0));//形态学膨胀

            return mat_dilate;
        }

        /// <summary>
        /// 筛选图中符合给定条件的轮廓
        /// </summary>
        /// <param name="mat_threshold">要提取轮廓的图片</param>
        /// <param name="width1">轮廓外接矩形大于该宽度值</param>
        /// <param name="width2">轮廓外接矩形小于该宽度值</param>
        /// <param name="height1">轮廓外接矩形大于该高度值</param>
        /// <param name="height2">轮廓外接矩形小于该高度值</param>
        public VectorOfVectorOfPoint GetUsefulContours(Mat mat, double ratio)
        {
            VectorOfVectorOfPoint contours = new VectorOfVectorOfPoint();//所有的轮廓
            VectorOfVectorOfPoint selected_contours = new VectorOfVectorOfPoint();//用于存储筛选过后的轮廓
            CvInvoke.FindContours(mat, contours, null, Emgu.CV.CvEnum.RetrType.List,
                Emgu.CV.CvEnum.ChainApproxMethod.ChainApproxSimple);//提取所有轮廓，操作过程中会对输入图像进行修改

            //筛选轮廓。筛选条件：长宽比大于给定值
            for (int i = 0; i < contours.Size; i++)
            {
                Rectangle rect = CvInvoke.BoundingRectangle(contours[i]);//外接矩形
                Mat temp = new Mat(mat, rect);//提取ROI矩形区域
                int pxNums = CvInvoke.CountNonZero(temp);//计算图像内非零像素个数

                double area = CvInvoke.ContourArea(contours[i]);//计算连通轮廓的面积
                double length = CvInvoke.ArcLength(contours[i], false); //计算连通轮廓的周长

                VectorOfPoint approx_curve = new VectorOfPoint();//用于存放逼近的结果
                CvInvoke.ApproxPolyDP(contours[i], approx_curve, length * 0.05, true);

                //外接矩形宽、高需在给定范围内
                bool bo = (rect.Width / rect.Height >= ratio && (rect.Width > 18 && rect.Height > 5) && area < 300);
                //bool bo = pxNums > 130 && pxNums < 300;
                if (bo)
                {
                    selected_contours.Push(contours[i]);
                }
            }
            return contours;
            //return selected_contours;
        }

        /// <summary>
        /// 根据给定的轮廓及范围信息计算涂卡的结果，返回int数组
        /// </summary>
        /// <param name="contours"></param>
        /// <param name="x"></param>
        /// <param name="x_interval"></param>
        /// <param name="x_num"></param>
        /// <param name="y"></param>
        /// <param name="y_interval"></param>
        /// <param name="y_num"></param>
        /// <returns></returns>
        public int[] GetTargetValues(VectorOfVectorOfPoint contours, int x_begin, int x_interval, int x_num,
            int y_begin, int y_interval, int y_num)
        {
            int[] result = new int[x_num];//结果数组
            //数组初值默认为-1
            for (int i = 0; i < x_num; i++)
            {
                result[i] = -1;
            }
            int x_max = x_begin + x_interval * x_num;
            int y_max = y_begin + y_interval * y_num;
            VectorOfVectorOfPoint targetContours = new VectorOfVectorOfPoint();
            Point[] gravity = GetGravityOfContours(contours);//轮廓中心点坐标

            for (int i = 0; i < contours.Size; i++)
            {
                VectorOfPoint contour = contours[i];
                if (gravity[i].X < x_begin || gravity[i].X > x_max || gravity[i].Y < y_begin || gravity[i].Y > y_max)
                {
                    continue;//判断中心点是否超出范围
                }
                int x_id = (int)Math.Floor((double)(gravity[i].X - x_begin) / x_interval);//向下取整
                int value = (int)Math.Floor((double)(gravity[i].Y - y_begin) / y_interval);

                if (result[x_id] != -1)
                {
                    string str = string.Format("第{0}列存在多个答案！请擦拭干净后再扫描", x_id);
                    MessageBox.Show(str);
                }
                else
                    result[x_id] = value;
            }
            return result;
        }

        /// <summary>
        /// 画出网格并返回填图结果
        /// </summary>
        /// <param name="img"></param>
        /// <param name="contours"></param>
        /// <param name="x_begin"></param>
        /// <param name="x_interval"></param>
        /// <param name="x_num"></param>
        /// <param name="y_begin"></param>
        /// <param name="y_interval"></param>
        /// <param name="y_num"></param>
        /// <param name="strText"></param>
        /// <returns></returns>
        public string GetValueAndDrawGrid(ImageBox img, VectorOfVectorOfPoint contours,
            int x_begin, int x_interval, int x_num, int y_begin, int y_interval, int y_num, string strText)
        {
            //画网格
            Mat src = new Image<Bgr, byte>(img.Image.Bitmap).Mat;
            Mat mat_grid = DrawGridByXY(img, x_begin, x_interval, x_num, y_begin, y_interval, y_num);

            int[] intArray = GetTargetValues(contours, x_begin, x_interval, x_num, y_begin, y_interval, y_num);
            int maxValue = GetMaxValueOfArray(intArray);//数组最大值

            string str = "";
            str += Environment.NewLine;//回车
            str += strText;
            if (maxValue >= 4)
            {
                str += GetStringOfIntArray(intArray);
            }
            else
            {
                str += GetStringOfIntArray(intArray, "ABCD");
            }

            return str;
        }

        /// <summary>
        /// 计算轮廓中心点坐标
        /// </summary>
        /// <param name="selected_contours">要计算中心点的轮廓</param>
        /// <returns></returns>
        public Point[] GetGravityOfContours(VectorOfVectorOfPoint selected_contours)
        {
            int ksize = selected_contours.Size;

            double[] m00 = new double[ksize];
            double[] m01 = new double[ksize];
            double[] m10 = new double[ksize];
            Point[] gravity = new Point[ksize];//用于存储轮廓中心点坐标
            MCvMoments[] moments = new MCvMoments[ksize];

            for (int i = 0; i < ksize; i++)
            {
                VectorOfPoint contour = selected_contours[i];
                //计算当前轮廓的矩
                moments[i] = CvInvoke.Moments(contour, false);

                m00[i] = moments[i].M00;
                m01[i] = moments[i].M01;
                m10[i] = moments[i].M10;
                int x = Convert.ToInt32(m10[i] / m00[i]);//计算当前轮廓中心点坐标
                int y = Convert.ToInt32(m01[i] / m00[i]);
                gravity[i] = new Point(x, y);
            }
            return gravity;
        }
        /// <summary>
        /// 根据给定XY初始坐标、间距、数量画网格
        /// </summary>
        /// <param name="mat"></param>
        /// <param name="x"></param>
        /// <param name="x_interval"></param>
        /// <param name="x_num"></param>
        /// <param name="y"></param>
        /// <param name="y_interval"></param>
        /// <param name="y_num"></param>
        /// <returns></returns>
        public Mat DrawGridByXY(ImageBox img, int x_begin, int x_interval, int x_num, int y_begin, int y_interval, int y_num)
        {
            Mat src = new Image<Bgr, byte>(img.Image.Bitmap).Mat;

            //转换颜色空间
            Mat mat_color = new Mat();
            if (src.NumberOfChannels == 1)
                CvInvoke.CvtColor(src, mat_color, Emgu.CV.CvEnum.ColorConversion.Gray2Bgr);
            else
                mat_color = src;

            for (int i = 0; i <= x_num; i++)
            {
                //先画竖线
                Point p1 = new Point(x_begin + x_interval * i, y_begin);
                Point p2 = new Point(x_begin + x_interval * i, y_begin + y_interval * y_num);
                CvInvoke.Line(mat_color, p1, p2, new MCvScalar(0, 0, 255), 1);
            }

            for (int i = 0; i <= y_num; i++)
            {
                //再画横线
                Point p1 = new Point(x_begin, y_begin + y_interval * i);
                Point p2 = new Point(x_begin + x_interval * x_num, y_begin + y_interval * i);
                CvInvoke.Line(mat_color, p1, p2, new MCvScalar(0, 0, 255), 1);
            }

            img.Image = mat_color;
            return mat_color;
        }

        /// <summary>
        /// 画出给定轮廓
        /// </summary>
        /// <param name="mat"></param>
        /// <param name="contours"></param>
        /// <returns></returns>
        public Mat DrawContours(Mat mat, VectorOfVectorOfPoint contours)
        {
            //转换颜色空间
            Mat mat_color = new Mat();
            CvInvoke.CvtColor(mat, mat_color, Emgu.CV.CvEnum.ColorConversion.Gray2Bgr);

            CvInvoke.DrawContours(mat_color, contours, -1, new MCvScalar(255, 0, 0), 2);
            return mat_color;
        }

        /// <summary>
        /// 拼接给定int数组内容，并返回拼接后字符串
        /// </summary>
        /// <param name="arr"></param>
        /// <returns></returns>
        public string GetStringOfIntArray(int[] arr)
        {
            string str = "";
            foreach (int a in arr)
            {
                str += a.ToString() + " ";
            }
            return str;
        }

        /// <summary>
        /// 拼接给定int数组内容，并返回拼接后字符串
        /// </summary>
        /// <param name="arr"></param>
        /// <returns></returns>
        public string GetStringOfIntArray(int[] arr, string ss = "ABCD")
        {
            string str = "";
            //char[] ch = new char[] { 'A', 'B', 'C', 'D', 'E', 'F', 'G' };
            char[] ch = ss.ToCharArray();
            foreach (int a in arr)
            {
                if (a == -1)
                {
                    str += "null ";//未识别时，标示为空
                }
                else
                {
                    str += ch[a] + " ";
                }

            }
            return str;
        }

        /// <summary>
        /// 获取一维int数组中最大值
        /// </summary>
        /// <param name="arr"></param>
        /// <returns></returns>
        public int GetMaxValueOfArray(int[] arr)
        {
            int[] dst = new int[arr.Length];
            Array.Copy(arr, dst, arr.Length);//深度复制数组，防止排序对原数组产生影响
            Array.Sort(dst);//数组排序
            int maxValue = dst[arr.Length - 1];//数组最大值
            return maxValue;
        }
    }
}
