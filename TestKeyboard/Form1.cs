using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using TestKeyboard.ClickKey;
using TestKeyboard.Entity;
using TestKeyboard.MonitorEvent;
using TestKeyboard.PressKey;
using TestKeyboard.Screen;
using TestKeyboard.SetForegroundWindow;
using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.Structure;
using Emgu.CV.Util;
using Emgu.CV.Flann;
using Emgu.CV.XFeatures2D;
using Emgu.CV.Features2D;       //包含Features2DToolbox 
using System.Text;
using System.Security.AccessControl;
using TestKeyboard.Picture;
using Emgu.CV.ML;
using TestKeyboard.DriverStageHelper;
using System.Collections.Concurrent;
using System.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using System.Net.Http;

namespace TestKeyboard
{
    public partial class Form1 : Form
    {
        private IPressKey mPressKey;
        private Foreground mFocus;
        private PressClick mClick;
        private bool isAdd = false;
        /// <summary>
        /// 键盘钩子KeyDown事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void HookMain_OnKeyDown(object sender, KeyEventArgs e)
        {
            SetText("键盘输入" + e.KeyValue);
            if (e.KeyCode == Keys.F1)
            {

                isStaring = false;
                RemoveEvent();
            }
        }
        /// <summary>
        /// 鼠标钩子点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void m_HookMain_OnMouseActivity(object sender, MouseEventArgs e)
        {
            try
            {
                //SetText("鼠标点击" + e.Button);
                if (e.Button == MouseButtons.Middle)
                {
                    

                    isStaring = false;
                    SaveConfig();
                    RemoveEvent();

                }
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }
        /// <summary>
        /// 钩子管理实例
        /// </summary>
        private SKHook m_HookMain = new SKHook();

        public Form1()
        {
            InitializeComponent();
            
            Control.CheckForIllegalCrossThreadCalls = false;
        }

        public DateTime runTime = DateTime.MinValue;//运行时间 
        public bool isStaring = false;//开启标志
        public List<job> jobs = new List<job>();
        BindingSource source;
        //任务类型
        Dictionary<int, jobType> dict = new Dictionary<int, jobType>();
        //键盘类型
        Dictionary<int, KeyBoard> dicBoard = new Dictionary<int, KeyBoard>();
        //鼠标类型
        Dictionary<int, EnumMouse> dicMouse = new Dictionary<int, EnumMouse>();
        //任务列表
        Dictionary<string, string> dicjobs = new Dictionary<string, string>();
        private void Form1_Load(object sender, EventArgs e)
        {
            runTime = DateTime.Now;
            
            //任务类型
            foreach (jobType item in Enum.GetValues(typeof(jobType)))
            { 
                dict.Add(item.GetHashCode(), item);
            } 
            comJob.DataSource = new BindingSource(dict, null);
            comJob.DisplayMember = "Value";
            comJob.ValueMember = "Key";

            //键盘类型 
            foreach (KeyBoard item in Enum.GetValues(typeof(KeyBoard)))
            {
                dicBoard.Add(item.GetHashCode(), item);
            } 
            comBoard.DataSource = new BindingSource(dicBoard, null);
            comBoard.DisplayMember = "Value";
            comBoard.ValueMember = "Key";


            //鼠标类型 
            foreach (EnumMouse item in Enum.GetValues(typeof(EnumMouse)))
            {
                dicMouse.Add(item.GetHashCode(), item);
            }
            comMouse.DataSource = new BindingSource(dicMouse, null);
            comMouse.DisplayMember = "Value";
            comMouse.ValueMember = "Key";

            //任务列表
            ReadTask();

            //数据源
            gridJobs.AutoGenerateColumns = true;
            gridJobs.RowHeadersVisible = true;
            gridJobs.AllowUserToAddRows = false;
            gridJobs.RowPostPaint += GridJobs_RowPostPaint;
            source = new BindingSource();
            source.DataSource = jobs; 
            gridJobs.DataSource = source;
            source.ResetBindings(false); 
            

           
            
        }
        public void ReadTask()
        {
            dicjobs.Clear();
            var dic = Path.Combine(Application.StartupPath, "jobs");
            if (!Directory.Exists(dic)) Directory.CreateDirectory(dic);
            var files = Directory.GetFiles(dic, "*.config");
            foreach (var item in files)
            {
                dicjobs.Add(Path.GetFileNameWithoutExtension(item), item);
            }
            comTask.DataSource = new BindingSource(dicjobs, null);
            comTask.DisplayMember = "Key";
            comTask.ValueMember = "Value";
        }
        private void GridJobs_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            Rectangle rectangle = new Rectangle(e.RowBounds.Location.X,
                e.RowBounds.Location.Y,
                gridJobs.RowHeadersWidth - 4,
                e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(),
                gridJobs.RowHeadersDefaultCellStyle.Font,
                rectangle,
                gridJobs.RowHeadersDefaultCellStyle.ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Right); 
        }
        ConcurrentBag<string> list = new ConcurrentBag<string>();
        public void SetText(string logStr,bool isClean=false)
        {
            try
            {
               

                Action log = new Action(() => { });
                log.BeginInvoke(ShowLog, logStr);
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
            
        }
        public void ShowLog(IAsyncResult ar)
        {
            try
            {
                //while (isStaring)
                //{
                //    richTextBox1.BeginInvoke(new Action(() => { richTextBox1.Text = sb.ToString(); }));
                //    Thread.Sleep(1000);
                //} 
                string logStr = ar.AsyncState.ToString();
                if (list.Count > 20)
                {
                    for (int i = 0; i < 30; i++)
                    {
                        string temp;
                        list.TryTake(out temp);
                    }
                }
                list.Add(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:ffff") + "-" + logStr);
                richLogs.Text = string.Join("\r\n", list.OrderByDescending(t => t));
                

            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }
        public void WriteTask(IAsyncResult ar)
        {
            try
            {
                while (isStaring)
                {
                    Thread.Sleep(30000);
                    SaveConfig();
                }
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        //SendInput
        private void button3_Click(object sender, EventArgs e)
        {
            mPressKey = new PressKeyBySendInput();
            bool bInitResult = mPressKey.Initialize(EnumWindowsType.Win64);
            if (bInitResult == false)
            {
                SetText("组件初始化失败", true);
            }
            richLogs.Focus();
            Thread.Sleep(1000);
            mPressKey.KeyPress('A');
            mPressKey.KeyPress('B');
            mPressKey.KeyPress('C');
        }
        //WinIO
        private void Button1_Click(object sender, EventArgs e)
        {
            mPressKey = new PressKeyByWinIO();
            bool bInitResult = mPressKey.Initialize(EnumWindowsType.Win64);
            if (bInitResult == false)
            {
                SetText("组件初始化失败");
                return;
            }
            richLogs.Focus();
            Thread.Sleep(1000);
            mPressKey.KeyPress('A');
            mPressKey.KeyPress('B');
            mPressKey.KeyPress('C');
        }
        //WinRing0
        private void button8_Click(object sender, EventArgs e)
        {
            mPressKey = new PressKeyByWinRing0();
            bool bInitResult = mPressKey.Initialize(EnumWindowsType.Win64);
            if (bInitResult == false)
            {
                SetText("组件初始化失败");
                return;
            }
            string msg = "";
            Foreground win = new Foreground();
            win.FocusWindow(windowName.Text, ref msg);
            Thread.Sleep(1000);
            mPressKey.KeyPress((char)KeyBoard.A);
            //mPressKey.KeyPress((char)KeyBoard.B);
            //mPressKey.KeyDown((char)KeyBoard.C);
            Thread.Sleep(100);
            mPressKey.KeyDown((char)KeyBoard.LeftArrow);
            mPressKey.KeyDown((char)KeyBoard.LeftArrow);
            Thread.Sleep(2000);
            mPressKey.KeyUp((char)KeyBoard.LeftArrow);
            mPressKey.KeyUp((char)KeyBoard.LeftArrow);
            Thread.Sleep(100);
            mPressKey.KeyDown((char)KeyBoard.RightArrow);
            mPressKey.KeyDown((char)KeyBoard.RightArrow);
            Thread.Sleep(2000);
            mPressKey.KeyUp((char)KeyBoard.RightArrow);
            mPressKey.KeyUp((char)KeyBoard.RightArrow);
        }
        bool isFirst = false;
        private void Button2_Click(object sender, EventArgs e)
        {
            AddEvent();
            Start();
        }
        public void Start()
        {
            try
            {
                SetText("任务开始",true);
                isFirst = true;
                if (isStaring)
                {
                    SetText("任务已开启,无需重复开启");
                    return;
                } 
                isStaring = true;
                mPressKey = new PressKeyByWinRing0();//按键
                mFocus = new Foreground();//聚焦
                mClick = new PressClick();//点击
                bool bInitResult = mPressKey.Initialize(EnumWindowsType.Win64);
                if (bInitResult == false)
                {
                    SetText("组件初始化失败");
                    return;
                }
                else
                {

                    //时间统计
                    Action calc = new Action(() => { });
                    calc.BeginInvoke(CaclTime, null);

                    //数据刷新
                    Action refresh = new Action(() => { });
                    calc.BeginInvoke(CalcData, null);

                    //定时回写
                    Action writetask = new Action(() => { });
                    writetask.BeginInvoke(WriteTask, null);

                    //定时任务
                    Action runjob = new Action(() => { });
                    runjob.BeginInvoke(RunJob, null);



                }
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

       
        public void RunJob(IAsyncResult ar)
        {
            try
            {


                SetText("任务开启");
                System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
                
                //循环任务
                while (isStaring)
                {
                    stopwatch.Reset();
                    stopwatch.Start();
                    var lessJobs = jobs.Where(t => t.count > 0 && t.countLess >= t.count).ToList();
                    if(lessJobs.Count == jobs.Count)
                    {
                        isStaring = false;
                        SetText("所有任务执行结束");
                        RemoveEvent();
                        break;
                    }
                    for (int i = 0; i < jobs.Count; i++)
                    {
                        if (!isStaring)
                            break;
                        job item = jobs[i];
                        if (item.count > 0 && item.countLess >= item.count)
                            continue;
                        item.typeName = item.type.ToString();
                        
                        Thread.Sleep(10);
                        string msg = "";

                        if (item.less <= 0 || item.less == item.time)
                        {
                            if (item.type == jobType.按键任务)
                            {
                                var key = (KeyBoard)item.content;
                                if (key == KeyBoard.LeftArrow || key == KeyBoard.RightArrow || key == KeyBoard.UpArrow || key == KeyBoard.DownArrow)
                                {
                                    var numlook = SendInputHelper.GetKeyState(144);
                                    
                                    if (numlook == 1)
                                    {
                                        mPressKey.KeyPress((char)KeyBoard.NumLock);//当按方向键时关闭小键盘
                                    }
                                    Thread.Sleep(10);
                                    mPressKey.KeyDown((char)(key == KeyBoard.LeftArrow ? KeyBoard.LeftArrow : KeyBoard.RightArrow));
                                    mPressKey.KeyDown((char)(key == KeyBoard.LeftArrow ? KeyBoard.LeftArrow : KeyBoard.RightArrow));
                                    Thread.Sleep((int)item.delay);
                                    mPressKey.KeyUp((char)(key == KeyBoard.LeftArrow ? KeyBoard.LeftArrow : KeyBoard.RightArrow));
                                    mPressKey.KeyUp((char)(key == KeyBoard.LeftArrow ? KeyBoard.LeftArrow : KeyBoard.RightArrow));
                                    Thread.Sleep(10);
                                    mPressKey.KeyDown((char)(key == KeyBoard.LeftArrow ? KeyBoard.RightArrow : KeyBoard.LeftArrow));
                                    mPressKey.KeyDown((char)(key == KeyBoard.LeftArrow ? KeyBoard.RightArrow : KeyBoard.LeftArrow));
                                    Thread.Sleep((int)item.delay);
                                    mPressKey.KeyUp((char)(key == KeyBoard.LeftArrow ? KeyBoard.RightArrow : KeyBoard.LeftArrow));
                                    mPressKey.KeyUp((char)(key == KeyBoard.LeftArrow ? KeyBoard.RightArrow : KeyBoard.LeftArrow));
                                    Thread.Sleep(10);
                                    if (numlook == 1)
                                    {
                                        mPressKey.KeyPress((char)KeyBoard.NumLock);//复原
                                    }
                                    //执行后左右取反
                                    if (key == KeyBoard.LeftArrow)
                                        item.content = KeyBoard.RightArrow;
                                    if (key == KeyBoard.RightArrow)
                                        item.content = KeyBoard.LeftArrow;
                                }else if (key == KeyBoard.截图)
                                {
                                    var screen = ScreenCat.GetScreen(checkScreen.Checked, (int)posX.Value, (int)posY.Value, (int)numWidth.Value, (int)numHeight.Value);
                                    var dic = Path.Combine(Application.StartupPath, "pic");
                                    if (!Directory.Exists(dic))
                                        Directory.CreateDirectory(dic);
                                    var savePath = Path.Combine(dic, $"{DateTime.Now.ToString("yyyMMddHHmmss")}.jpg");
                                    ScreenCat.SaveImageWithQuality(savePath, screen, 100);
                                    pictureBox1.Image = screen;
                                    SetText($"截图成功:{savePath}");
                                }
                                else
                                {
                                    mPressKey.KeyPress((char)key);
                                }

                                SetText("按键" + item.content.ToString());
                            }
                            else if (item.type == jobType.点击任务)
                            {

                                mClick.Click(item.x, item.y, (EnumMouse)item.content);
                                SetText("点击" + item.content.ToString());
                            }
                            else if (item.type == jobType.聚焦窗体)
                            {
                                var isOk = mFocus.FocusWindow(item.content.ToString(), ref msg);
                                SetText(msg);
                                if (!isOk)
                                {
                                    SetText("任务主动停止");
                                    isStaring = false;
                                }
                                else
                                {
                                    SetText("聚焦" + item.content.ToString());
                                }
                            }
                            else if (item.type == jobType.移动窗体)
                            {
                                var isOk = ScreenCat.Move(item.content.ToString());
                                if (!isOk)
                                {
                                    SetText("窗口移动失败");
                                    isStaring = false;
                                }
                                else
                                {
                                    SetText("移动" + item.content.ToString());
                                }
                            }
                            //后置延迟
                            bool isDely = true;
                            if (item.type == jobType.按键任务)
                            {
                                var key = (KeyBoard)item.content;
                                if (key == KeyBoard.LeftArrow || key == KeyBoard.RightArrow || key == KeyBoard.UpArrow || key == KeyBoard.DownArrow)
                                {
                                    isDely = false;
                                }
                            }
                            if (isDely)
                                Thread.Sleep((int)item.delay);


                            //stopwatch.Stop();
                            //if (item.less == item.time)
                            //    item.less -= stopwatch.ElapsedMilliseconds;
                            //if (item.less <= 0)
                            //    item.less = item.time - stopwatch.ElapsedMilliseconds;
                            item.countLess += 1;//执行记录

                        }
                        else
                        {
                            //stopwatch.Stop();
                            //item.less -= stopwatch.ElapsedMilliseconds;
                        }

                       
                    }
                    stopwatch.Stop();
                    foreach (var other in jobs)
                    {
                        other.less -= stopwatch.ElapsedMilliseconds;
                        if (other.less < 0) other.less = other.time;
                    }
                    isFirst = false;
                }
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
                Thread.Sleep(1000);
                isStaring = false;
            }
        }
        public void CalcData(IAsyncResult ar)
        {
            //刷新
            try
            {
                SetText("开启刷新");
                while (isStaring)
                {
                    RefreshData();
                    Thread.Sleep(1000);
                }
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }

        }
        public void CaclTime(IAsyncResult ar)
        {
            //统计
            try
            {
                SetText("开启计时");
                while (true)
                {
                    txtSecond.BeginInvoke(new Action(() =>
                    {
                        int total = (int)(DateTime.Now - runTime).TotalSeconds;
                        txtHour.Text = (total / 3600).ToString();
                        txtMin.Text = ((total - (total / 3600 * 3600)) / 60).ToString();
                        txtSecond.Text = (total % 60).ToString();
                    }));
                    Thread.Sleep(1000);
                }
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                isStaring = false;
                SaveConfig();
                RemoveEvent();
                SetText("任务停止");
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }

        }


        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                Foreground win = new Foreground();
                string msg = "";
                win.FocusWindow(windowName.Text,ref msg);
                SetText(msg);
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
           
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                Action action = new Action(() => { });
                action.BeginInvoke(TestPress, null);

            }
            catch (Exception ex)
            {

                SetText(ex.Message);
            }
        }
        public void TestPress(IAsyncResult data)
        {
            mPressKey = new PressKeyByWinRing0(); 
            bool bInitResult = mPressKey.Initialize(EnumWindowsType.Win64);
            Foreground win = new Foreground();
            string msg = "";
            if (bInitResult == false)
            {
                SetText("组件初始化失败", true);
                return;
            }
            else
            {

                win.FocusWindow(windowName.Text, ref msg);
                SetText(msg);
                Thread.Sleep(100); 
                mPressKey.KeyPress((char)KeyBoard.D1);
                SetText("按下");  
            }
        } 

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                Action action = new Action(() => { });
                action.BeginInvoke(TestClick,null);
                
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        } 
        public void TestClick(IAsyncResult data)
        {
            PressClick pressClick = new PressClick();
            Thread.Sleep(1000);
            pressClick.Click((int)posX.Value, (int)posY.Value, EnumMouse.鼠标右键);
            SetText("点击x:"+ (int)posX.Value + "|y:"+ (int)posY.Value);
        }
        KeyEventHandler keyEvent = null;
        MouseUpdateEventHandler mouseUpdateEvent = null;
        MouseEventHandler mouseEvent = null;
        public void AddEvent()
        {
            //keyEvent = new KeyEventHandler(HookMain_OnKeyDown);
            //mouseUpdateEvent = new MouseUpdateEventHandler(HookMain_OnMouseUpdate);
            mouseEvent = new MouseEventHandler(m_HookMain_OnMouseActivity);
            //m_HookMain.OnKeyDown += keyEvent;
            //m_HookMain.OnMouseUpdate += mouseUpdateEvent;
            SetText("开启鼠标键盘监听");
            m_HookMain.OnMouseActivity += mouseEvent;
            m_HookMain.InstallHook();
            
        }
        public void RemoveEvent()
        { 
            m_HookMain.OnKeyDown -= keyEvent;
            m_HookMain.OnMouseUpdate -= mouseUpdateEvent;
            m_HookMain.OnMouseActivity -= mouseEvent;
            m_HookMain.UnInstallHook();
            SetText("关闭鼠标键盘监听");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                isStaring = true;
                AddEvent();
                Action action = new Action(() => { });
                action.BeginInvoke(GetClickPos, null);
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }
        public void GetClickPos(IAsyncResult data)
        {
            PressClick pressClick = new PressClick();
            while (isStaring)
            {
                var pos = pressClick.GetClickPos();
                if (pos != null)
                {
                    posX.BeginInvoke(new Action(() => { posX.Value = pos.X; posY.Value = pos.Y; }));
                }
                Thread.Sleep(100);
            }
            
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            try
            { 
                var jtype = (KeyValuePair<int, jobType>)comJob.SelectedItem;
                job job = new job();
                if (jtype.Key == jobType.按键任务.GetHashCode())
                {
                    job = new job
                    {
                        type = (jobType)this.comJob.SelectedValue,
                        time = Convert.ToDecimal(numTime.Text) * 1000,
                        less = Convert.ToDecimal(numTime.Text) * 1000,
                        content = (KeyBoard)this.comBoard.SelectedValue,
                        delay = Convert.ToDecimal(numDelay.Text) * 1000

                    };
                }
                else if (jtype.Key == jobType.点击任务.GetHashCode())
                {
                    job = new job
                    {
                        type = (jobType)this.comJob.SelectedValue,
                        time = Convert.ToDecimal(numTime.Text) * 1000,
                        less = Convert.ToDecimal(numTime.Text) * 1000,
                        content = (EnumMouse)this.comMouse.SelectedValue,
                        delay = Convert.ToDecimal(numDelay.Text) * 1000,
                        x = (int)posX.Value,
                        y = (int)posY.Value
                    };
                }
                else if (jtype.Key == jobType.聚焦窗体.GetHashCode())
                {
                    job = new job
                    {
                        type = (jobType)this.comJob.SelectedValue,
                        time = Convert.ToDecimal(numTime.Text) * 1000,
                        less = Convert.ToDecimal(numTime.Text) * 1000,
                        content = this.windowName.Text,
                        delay = Convert.ToDecimal(numDelay.Text) * 1000
                    };
                }
                else if (jtype.Key == jobType.移动窗体.GetHashCode())
                {
                    job = new job
                    {
                        type = (jobType)this.comJob.SelectedValue,
                        time = Convert.ToDecimal(numTime.Text) * 1000,
                        less = Convert.ToDecimal(numTime.Text) * 1000,
                        content = this.windowName.Text,
                        x = Convert.ToInt32(posX.Value),
                        y = Convert.ToInt32(posY.Value),
                        delay = Convert.ToDecimal(numDelay.Text) * 1000
                    };
                   
                }
                jobs.Add(job);
                job.contentName = job.content.ToString();
                job.typeName = job.type.ToString();
                job.count = Convert.ToInt32(numCount.Value);


                RefreshData();
                SetText("添加成功");
                SaveConfig();
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }
        public void SaveConfig()
        {
            var config = Newtonsoft.Json.JsonConvert.SerializeObject(jobs, Formatting.Indented);
            File.WriteAllText(comTask.SelectedValue.ToString(), config); 
        }
        public void ReadConfig()
        {
            jobs.Clear();
            if (File.Exists(comTask.SelectedValue.ToString()))
            {
                var config = File.ReadAllText(comTask.SelectedValue.ToString());
                List<job> ls = Newtonsoft.Json.JsonConvert.DeserializeObject<List<job>>(config);
                if(ls != null && ls.Count > 0)
                {
                    foreach (var item in ls)
                    {
                        if(item.type == jobType.按键任务 && item.content is long)
                            item.content = (KeyBoard)Enum.ToObject(typeof(KeyBoard), Convert.ToInt32(item.content));
                        if (item.type == jobType.点击任务 && item.content is long)
                            item.content = (EnumMouse)Enum.ToObject(typeof(EnumMouse), Convert.ToInt32(item.content));
                    }
                    jobs.AddRange(ls);
                }
            }
            RefreshData();

        }
        private void btnDel_Click(object sender, EventArgs e)
        {
            try
            {
                var row = gridJobs.SelectedRows;
                if (row.Count > 0)
                {
                    List<job> ls = new List<job>();
                    foreach (DataGridViewRow item in row)
                    {
                        var tjob = (job)item.DataBoundItem;
                        ls.Add(tjob); 
                    }
                    foreach (var item in ls)
                    {
                        jobs.Remove(item);
                    }
                    
                    RefreshData(); 
                    SaveConfig();
                    SetText("删除成功");
                }
                else
                {
                    SetText("请选中要删除的行");
                }
               

            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }
        /// <summary>
        /// 刷新表格
        /// </summary>
        public void RefreshData(IAsyncResult data)
        {
            RefreshData();
        }
        /// <summary>
        /// 刷新表格
        /// </summary>
        public void RefreshData()
        {
           
            this.gridJobs.BeginInvoke(new Action(() =>
            {
                source.ResetBindings(false);
            }));
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            Reset();
        }
        public void Reset()
        {
            try
            {
                foreach (var item in jobs)
                {
                    item.less = item.time;
                    item.countLess = 0;
                }
                RefreshData();
                SaveConfig();
                SetText($"任务重置成功");

            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Action action = new Action(() => { });
            action.BeginInvoke(TestScreen, null);
        }
        public void TestScreen(IAsyncResult data)
        {
            try
            {
                Thread.Sleep(2000);
                var screen = ScreenCat.GetScreen(checkScreen.Checked, (int)posX.Value, (int)posY.Value, (int)numWidth.Value, (int)numHeight.Value);
                var savePath = Path.Combine(Application.StartupPath, $"{DateTime.Now.ToString("yyyMMddHHmmss")}.jpg");
                ScreenCat.SaveImageWithQuality(savePath, screen, 100);
                pictureBox1.Image = screen;
                SetText($"截图成功:{savePath}");
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            { 
                String imagePath = Path.Combine(Application.StartupPath, txtpos.Text);
                String imagePath2 = Path.Combine(Application.StartupPath, txtneg.Text); 
                
                switch ("1")
                {
                    case "1":
                        var data = GetMatchPos(imagePath, imagePath2);
                        PressClick pressClick = new PressClick();
                        pressClick.Click(data.X, data.Y, EnumMouse.鼠标右键);
                        SetText("点击x:" + data.X + "|y:" + data.Y);
                        break;
                    case "2":
                        GetMatchPos2(imagePath, imagePath2);
                        break;
                    case "3":
                        GetMatchPos3(imagePath, imagePath2);
                        break;
                    case "4":
                        GetMatchPos4(imagePath, imagePath2);
                        break;
                    case "5":
                        GetMatchPos5(imagePath, imagePath2);
                        break;
                    case "6":
                        GetMatchPos6(imagePath, imagePath2);
                        break;
                    case "7":
                        GetMatchPos7(imagePath, imagePath2);
                        break;
                    case "8":
                        GetMatchPos8(imagePath, imagePath2);
                        break;
                    case "9":
                        GetMatchPos9(imagePath, imagePath2);
                        break;
                    default:
                        break;
                }
                

            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }
        public void GetMatchPos9(string bigImg, string smallImg)
        {
            ///亚像素角点检测
            Mat src = CvInvoke.Imread(bigImg, LoadImageType.Color);
            CvInvoke.Imshow("input", src);

            Mat grayImg = new Mat();
            CvInvoke.CvtColor(src, grayImg, ColorConversion.Bgr2Gray);
            Mat harrisImg = new Mat();
            CvInvoke.CornerHarris(grayImg, harrisImg, 2, 3, 0.04);
            CvInvoke.Threshold(harrisImg, harrisImg, 0.01, 255, ThresholdType.Binary);
            Mat dst = new Mat();
            harrisImg.ConvertTo(dst, DepthType.Cv8U);
            CvInvoke.Imshow("mask", dst);


            VectorOfPointF corners = new VectorOfPointF();
            Image<Gray, Byte> img = dst.ToImage<Gray, Byte>();
            for (int i = 0; i < img.Rows; i++)
            {
                for (int j = 0; j < img.Cols; j++)
                {
                    if (img.Data[i, j, 0] == 255)
                    {
                        CvInvoke.Circle(src, new Point(j, i), 3, new MCvScalar(0, 0, 255), -1);
                        PointF[] pt = new PointF[1];
                        pt[0].X = j;
                        pt[0].Y = i;
                        corners.Push(pt);
                    }
                }
            }
            CvInvoke.Imshow("result", src);
            MCvTermCriteria termCriteria = new MCvTermCriteria(40, 0.001);
            CvInvoke.CornerSubPix(grayImg, corners, new Size(5, 5), new Size(-1, -1), termCriteria); //亚像素级角点精确化
            for (int i = 0; i < corners.Size; i++)
            {
                Console.WriteLine("corner {0} : ({1:F2},{1:F2})", i, corners[i].X, corners[i].Y);
            }

            CvInvoke.WaitKey(0);
        }
        public void GetMatchPos8(string bigImg, string smallImg)
        {
            ///另一种方法
            Mat src = CvInvoke.Imread(bigImg, LoadImageType.Color);
            Mat grayImg = new Mat();
            CvInvoke.CvtColor(src, grayImg, ColorConversion.Bgr2Gray);

            Mat dst = new Mat();
            CvInvoke.CornerHarris(grayImg, dst, 2, 3, 0.04, BorderType.Default);
            Mat scaleImg = new Mat();
            CvInvoke.Normalize(dst, dst, 0, 255, NormType.MinMax, DepthType.Cv32F);
            CvInvoke.ConvertScaleAbs(dst, scaleImg, 1, 0);
            Image<Gray, Byte> img = scaleImg.ToImage<Gray, Byte>();
            for (int i = 0; i < img.Rows; i++)
            {
                for (int j = 0; j < img.Cols; j++)
                {
                    if (img.Data[i, j, 0] > 100)           //阈值选取很重要，控制角点个数
                    {
                        CvInvoke.Circle(src, new Point(j, i), 3, new MCvScalar(0, 255, 0), -1);
                    }
                }
            }
            CvInvoke.Imshow("result", src);
            CvInvoke.WaitKey(0);
        }
        public void GetMatchPos7(string bigImg, string smallImg)
        {
            //Harris角点检测
            Mat src = CvInvoke.Imread(bigImg, LoadImageType.Color);
            CvInvoke.Imshow("input", src);
            Mat gray = new Mat();
            CvInvoke.CvtColor(src, gray, ColorConversion.Bgr2Gray);
            Mat dst = new Mat();
            CvInvoke.CornerHarris(gray, dst, 2, 3, 0.04);//角点检测
            CvInvoke.Threshold(dst, dst, 0.005, 255, ThresholdType.Binary);
            Console.WriteLine("Depth: {0}\nChannels: {1}", dst.Depth, dst.NumberOfChannels);
            dst.ConvertTo(dst, DepthType.Cv8U);
            CvInvoke.Imshow("CornerHarris", dst);
            Image<Gray, Byte> img = dst.ToImage<Gray, Byte>();
            for (int i = 0; i < img.Rows; i++)
            {
                for (int j = 0; j < img.Cols; j++)
                {
                    if (img.Data[i, j, 0] == 255)
                    {
                        CvInvoke.Circle(src, new Point(j, i), 2, new MCvScalar(0, 255, 0), -1);
                    }
                }
            }
            CvInvoke.Imshow("result", src);
        }
        public void GetMatchPos6(string bigImg, string smallImg)
        {
                Mat src = new Image<Bgr, byte>(bigImg).Mat;
            CommonUse commonUse = new CommonUse();
                //1.获取当前图像的最大矩形边界
                VectorOfVectorOfPoint max_contour = commonUse.GetBoundaryOfPic(src);

                //2.对图像进行矫正
                Mat mat_Perspective = commonUse.MyWarpPerspective(src, max_contour);
                //规范图像大小
                //CvInvoke.Resize(mat_Perspective, mat_Perspective, new Size(590, 384), 0, 0, Emgu.CV.CvEnum.Inter.Cubic);
                //3.二值化处理（大于阈值取0，小于阈值取255。其中白色为0，黑色为255)
                Mat mat_threshold = new Mat();
                int myThreshold = Convert.ToInt32(0);
                CvInvoke.Threshold(mat_Perspective, mat_threshold, myThreshold, 255, Emgu.CV.CvEnum.ThresholdType.BinaryInv);
                //ib_middle.Image = mat_threshold;

                //形态学膨胀
                Mat mat_dilate = commonUse.MyDilate(mat_threshold);
                //ib_middle.Image = mat_dilate;

                //筛选长宽比大于2的轮廓
                VectorOfVectorOfPoint selected_contours = commonUse.GetUsefulContours(mat_dilate, 1);
                //画出轮廓
                Mat color_mat = commonUse.DrawContours(mat_Perspective, selected_contours);
                 CvInvoke.Imshow("box", color_mat);


                //ib_result.Image = mat_Perspective;
                ////准考证号，x=230+26*5,y=40+17*10
                //tb_log.Text = commonUse.GetValueAndDrawGrid(ib_result, selected_contours, 230, 26, 5, 40, 17, 10, "准考证号：");

                ////答题区1-5题，x=8+25*5,y=230+16*4
                //tb_log.Text += commonUse.GetValueAndDrawGrid(ib_result, selected_contours, 8, 25, 5, 230, 16, 4, "1-5：");

                ////答题区6-10题,x=159+25*5,y=230+16*4
                //tb_log.Text += commonUse.GetValueAndDrawGrid(ib_result, selected_contours, 159, 25, 5, 230, 16, 4, "6-10：");

                ////答题区11-15题,x=310+25*5,y=230+16*4
                //tb_log.Text += commonUse.GetValueAndDrawGrid(ib_result, selected_contours, 310, 25, 5, 230, 16, 4, "11-15：");

                ////答题区16-20题,x=461+25*5,y=230+16*4
                //tb_log.Text += commonUse.GetValueAndDrawGrid(ib_result, selected_contours, 461, 25, 5, 230, 16, 4, "16-20：");

                ////答题区21-25题,x=8+25*5,y=312+16*4
                //tb_log.Text += commonUse.GetValueAndDrawGrid(ib_result, selected_contours, 8, 25, 5, 312, 16, 4, "21-25：");

                ////答题区26-30题,x=159+25*5,y=312+16*4
                //tb_log.Text += commonUse.GetValueAndDrawGrid(ib_result, selected_contours, 159, 25, 5, 312, 16, 4, "26-30：");

                ////答题区31-35题,x=310+25*5,y=312+16*4
                //tb_log.Text += commonUse.GetValueAndDrawGrid(ib_result, selected_contours, 310, 25, 5, 312, 16, 4, "31-35：");
            
        }
        public void GetMatchPos5(string bigImg, string smallImg)
        {
            Mat readImage = CvInvoke.Imread(bigImg, LoadImageType.Color);
            Mat grayImage = new Mat();
            CvInvoke.CvtColor(readImage, grayImage, ColorConversion.Rgb2Gray);
            Mat blurImage = new Mat();
            CvInvoke.GaussianBlur(grayImage, blurImage, new Size(3, 3), 3);
            Mat cannyImage = new Mat();
            CvInvoke.Canny(blurImage, cannyImage, 60, 180);
            CvInvoke.Imshow("empty", cannyImage);

            List<Triangle2DF> triangleList = new List<Triangle2DF>();  //三角形列表
            List<RotatedRect> rectList = new List<RotatedRect>();//矩形列表
            //using(){}当我们做一些比较占用资源的操作,在此范围的末尾自动将对象释放。
            //VectorOfVectorOfPoint:创建一个空的VectorOfPoint标准向量
            using (VectorOfVectorOfPoint contours = new VectorOfVectorOfPoint())
            {
                //  参数：输入图像，点向量，，检索类型，方法
                CvInvoke.FindContours(cannyImage, contours, null, RetrType.List, ChainApproxMethod.ChainApproxSimple);
                for (int i = 0; i < contours.Size; i++)
                {
                    using (VectorOfPoint contour = contours[i])  //点的标准矢量
                    using (VectorOfPoint approxContour = new VectorOfPoint())
                    {
                        // CvInvoke.ApproxPolyDP：指定精度的多边形曲线  参数：输入向量（轮廓），近似轮廓，近似精度，true：则闭合形状
                        // CvInvoke.ArcLength：计算轮廓长度   参数：轮廓，曲线是否闭合
                        CvInvoke.ApproxPolyDP(contour, approxContour, CvInvoke.ArcLength(contour, true) * 0.08, true);  //0.08的相似度
                        //CvInvoke.ContourArea：计算轮廓面积  参数：轮廓，返回值是否为绝对值
                        //仅考虑面积大于50的轮廓
                        if (CvInvoke.ContourArea(approxContour, false) > 50)
                        {
                            if (approxContour.Size == 3)//轮廓有3个顶点：三角形
                            {
                                Point[] points = approxContour.ToArray(); //转化为数组
                                triangleList.Add(new Triangle2DF(points[0], points[1], points[2]));
                            }
                            else if (approxContour.Size == 4)//轮廓有4个顶点:矩形
                            {
                                bool isRect = true;
                                Point[] points = approxContour.ToArray();
                                LineSegment2D[] rects = PointCollection.PolyLine(points, true);
                                for (int j = 0; j < rects.Length; j++)
                                {
                                    double angle = Math.Abs(
                                        rects[(j + 1) % rects.Length].GetExteriorAngleDegree(rects[j]));
                                    //角度在小于80，大于100，判定不是矩形
                                    //if (angle < 80 || angle > 100)
                                    //{
                                    //    isRect = false;
                                    //    break;
                                    //}
                                }
                                if (isRect) { rectList.Add(CvInvoke.MinAreaRect(approxContour)); }
                            }
                        }
                    }
                }

            }

            //展示结果
            //三角形
            //Image<Bgr, Byte> image = readImage.ToImage<Bgr, Byte>().Copy();
            //foreach (Triangle2DF triangle in triangleList)
            //{
            //    image.Draw(triangle, new Bgr(Color.Red), 2);
            //}
            //CvInvoke.Imshow("Triangle Image", image);
            //矩形
            Image<Bgr, Byte> iamge1 = readImage.ToImage<Bgr, Byte>().Copy();
            foreach (RotatedRect rect in rectList)
            {
                iamge1.Draw(rect, new Bgr(Color.Red), 2);
            }
            CvInvoke.Imshow("Rect Image", iamge1);
            CvInvoke.WaitKey();
        }
        public void GetMatchPos4(string bigImg, string smallImg)
        {
            //加载原图
            var image1 = new Image<Bgr, byte>(bigImg);
            var image0 = image1.Mat.Clone(); 
            //CvInvoke.Imshow("1", image1);

            // 2. 灰度图
            var img2 = image0.Clone();
            CvInvoke.CvtColor(image0, img2, ColorConversion.Bgr2Gray);
            //CvInvoke.Imshow("2", img2);

            // 3. 阈值
            var img3 = new Mat();
            CvInvoke.Threshold(img2, img3, 0, 255, ThresholdType.Otsu);
            CvInvoke.Imshow("3", img3);

            // 4. 查找轮廓
            VectorOfVectorOfPoint contours = new VectorOfVectorOfPoint();
            Mat hierarchy = new Mat();
            CvInvoke.FindContours(img3, contours, hierarchy, RetrType.Tree, ChainApproxMethod.ChainApproxNone);

            // 5. 绘制所有轮廓
            var img4 = image0.Clone();
            CvInvoke.DrawContours(img4, contours, -1, new MCvScalar(0, 0, 255), 2);
            CvInvoke.Imshow("4", img4);

            // 6. 绘制最大的轮廓
            double MaxArea = 0;
            int maxIndex = 0;
            for (int i = 0; i < contours.Size; i++)
            {
                var area = CvInvoke.ContourArea(contours[i]);
                if (area > MaxArea)
                {
                    MaxArea = area;
                    maxIndex = i;
                }
            }
            var img5 = image0.Clone();
            CvInvoke.DrawContours(img5, contours, maxIndex, new MCvScalar(0, 0, 255), 2);
            //CvInvoke.Imshow("5", img5);

            // 7. 轮廓近似
            var e = CvInvoke.ArcLength(contours[maxIndex], true) * 0.01;
            VectorOfPoint s = new VectorOfPoint();
            CvInvoke.ApproxPolyDP(contours[maxIndex], s, e, true);
            var img6 = image0.Clone();
            VectorOfVectorOfPoint contours2 = new VectorOfVectorOfPoint();
            contours2.Push(s);
            CvInvoke.DrawContours(img6, contours2, -1, new MCvScalar(0, 0, 255), 2);
            //CvInvoke.Imshow("6", img6);

            // 8. 边界矩形
            var rect = CvInvoke.BoundingRectangle(contours[maxIndex]);
            var img7 = image0.Clone();
            CvInvoke.Rectangle(img7, rect, new MCvScalar(0, 0, 255), 2);
            //CvInvoke.Imshow("7", img7);

            // 9. 外接圆
            var img8 = image0.Clone();
            var circle = CvInvoke.MinEnclosingCircle(contours[maxIndex]);
            CvInvoke.Circle(img8, new Point((int)circle.Center.X, (int)circle.Center.Y), (int)circle.Radius, new MCvScalar(0, 0, 255), 2);
            //CvInvoke.Imshow("8", img8);

            // 10. 外接三角形
            var img9 = image0.Clone();
            VectorOfPoint triangle = new VectorOfPoint();
            CvInvoke.MinEnclosingTriangle(contours[maxIndex], triangle);
            CvInvoke.Line(img9, triangle[0], triangle[1], new MCvScalar(0, 0, 255), 2);
            CvInvoke.Line(img9, triangle[1], triangle[2], new MCvScalar(0, 0, 255), 2);
            CvInvoke.Line(img9, triangle[2], triangle[0], new MCvScalar(0, 0, 255), 2); 
            //CvInvoke.Imshow("9", img9);

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="img1">大图</param>
        /// <param name="img2">小图</param>
        /// <returns></returns>
        public Rectangle GetMatchPos(string img1, string img2)
        {
            var screen = ScreenCat.GetScreen();
            var savePath = Path.Combine(Application.StartupPath, "data.png");
            ScreenCat.SaveImageWithQuality(savePath, screen, 50);
            Mat Src = null;
            if (string.IsNullOrEmpty(img1))
                Src = CvInvoke.Imread(savePath, LoadImageType.Grayscale);
            else
                Src = CvInvoke.Imread(img1, LoadImageType.Grayscale);

            Mat Template = CvInvoke.Imread(img2, LoadImageType.Grayscale);

            Mat MatchResult = new Mat();//匹配结果
            CvInvoke.MatchTemplate(Src, Template, MatchResult, Emgu.CV.CvEnum.TemplateMatchingType.CcorrNormed);//使用相关系数法匹配
            Point max_loc = new Point();
            Point min_loc = new Point();
            double max = 0, min = 0;
            CvInvoke.MinMaxLoc(MatchResult, ref min, ref max, ref min_loc, ref max_loc);//获得极值信息

            var pos = new Rectangle(max_loc, Template.Size);
            SetText($"位置:x={pos.X},y={pos.Y}|宽度:width={pos.Width},高度={pos.Height}");
            return pos;
        }
        public void GetMatchPos2(string img1, string img2)
        {
            Mat src1 = CvInvoke.Imread(img1, LoadImageType.AnyDepth);
            Mat src2 = CvInvoke.Imread(img2, LoadImageType.AnyDepth);

            CvInvoke.Imshow("src1", src1);
            CvInvoke.Imshow("src2", src2);

            SURF surf = new SURF(400);
            //计算特征点
            MKeyPoint[] keyPoints1 = surf.Detect(src1);
            MKeyPoint[] keyPoints2 = surf.Detect(src2);
            VectorOfKeyPoint vkeyPoints1 = new VectorOfKeyPoint(keyPoints1);
            VectorOfKeyPoint vkeyPoints2 = new VectorOfKeyPoint(keyPoints2);
            Mat suft_feature1 = new Mat();
            Mat suft_feature2 = new Mat();
            //绘制特征点
            Features2DToolbox.DrawKeypoints(src1, vkeyPoints1, suft_feature1, new Bgr(0, 255, 0), Features2DToolbox.KeypointDrawType.Default);
            Features2DToolbox.DrawKeypoints(src2, vkeyPoints2, suft_feature2, new Bgr(0, 255, 0), Features2DToolbox.KeypointDrawType.Default);
            //显示特征点
            CvInvoke.Imshow("suft_feature1", suft_feature1);
            CvInvoke.Imshow("suft_feature2", suft_feature2);

            //计算特征描述符
            Mat descriptors1 = new Mat();
            Mat descriptors2 = new Mat();
            surf.Compute(src1, vkeyPoints1, descriptors1);
            surf.Compute(src2, vkeyPoints2, descriptors2);
            ///匹配方法1
            //设置暴力匹配器进行匹配显示
            //VectorOfVectorOfDMatch matches = new VectorOfVectorOfDMatch();
            //BFMatcher bFMatcher = new BFMatcher(DistanceType.L2);
            //bFMatcher.Add(descriptors1);
            //bFMatcher.KnnMatch(descriptors2, matches, 2, null);

            //筛选符合条件的匹配点
            //double min_dis = 100, max_dis = 0;
            //for (int i = 0; i < descriptors1.Rows; i++)
            //{
            //    不知道为啥会抛出异常
            //    try
            //    {
            //        if (max_dis < matches[i][0].Distance)
            //        {
            //            max_dis = matches[i][0].Distance;
            //        }
            //        if (min_dis > matches[i][0].Distance)
            //        {
            //            min_dis = matches[i][0].Distance;
            //        }
            //    }
            //    catch (Exception)
            //    {
            //        Console.WriteLine("exception...");
            //    }

            //}
            //VectorOfVectorOfDMatch good_matches = new VectorOfVectorOfDMatch();
            //for (int i = 0; i < matches.Size; i++)
            //{
            //    if (matches[i][0].Distance < 2 * min_dis)        //倍数关系自由调整
            //    {
            //        good_matches.Push(matches[i]);
            //    }
            //}
            //不使用掩膜，直接绘制出匹配的特征点
            //Mat result = new Mat();
            //Features2DToolbox.DrawMatches(src1, vkeyPoints1, src2, vkeyPoints2, good_matches, result,
            //                              new MCvScalar(0, 255, 0), new MCvScalar(0, 0, 255), null);
            //显示最终结果
            //CvInvoke.Imshow("match-result", result);
            //CvInvoke.WaitKey(0);

            ///匹配方法2-Flann快速匹配
            //创建Flann对象
            IIndexParams id = new LinearIndexParams();
            SearchParams sp = new SearchParams();
            //FlannBasedMatcher flannBasedMatcher = new FlannBasedMatcher(id, sp);
            //添加模板
            //flannBasedMatcher.Add(descriptors1);
            //计算匹配
            VectorOfVectorOfDMatch matches = new VectorOfVectorOfDMatch();
            //flannBasedMatcher.KnnMatch(descriptors2, matches, 2, null);

            Mat mask = new Mat(matches.Size, 1, DepthType.Cv8U, 1);
            mask.SetTo(new MCvScalar(255));
            Features2DToolbox.VoteForUniqueness(matches, 0.7, mask); //计算掩膜
            //绘制并显示匹配的特征点
            Mat result = new Mat();
            Features2DToolbox.DrawMatches(src1, vkeyPoints1, src2, vkeyPoints2, matches, result, new MCvScalar(0, 255, 0),
                new MCvScalar(0, 0, 255), mask);

            CvInvoke.Imshow("match-result", result);
            CvInvoke.WaitKey(0);
        }
        public void GetMatchPos3(string img1, string img2)
        {
            Mat src = CvInvoke.Imread(img1, LoadImageType.AnyColor);//从本地读取图片
            Mat result = src.Clone();

            Mat tempImg = CvInvoke.Imread(img2, LoadImageType.AnyColor);
            int matchImg_rows = src.Rows - tempImg.Rows + 1;
            int matchImg_cols = src.Cols - tempImg.Cols + 1;
            Mat matchImg = new Mat(matchImg_rows, matchImg_rows, DepthType.Cv32F, 1); //存储匹配结果
            #region 模板匹配参数说明
            ////采用系数匹配法，匹配值越大越接近准确图像。
            ////IInputArray image：输入待搜索的图像。图像类型为8位或32位浮点类型。设图像的大小为[W, H]。
            ////IInputArray templ：输入模板图像，类型与待搜索图像类型一致，并且大小不能大于待搜索图像。设图像大小为[w, h]。
            ////IOutputArray result：输出匹配的结果，单通道，32位浮点类型且大小为[W - w + 1, H - h + 1]。
            ////TemplateMatchingType method：枚举类型标识符，表示匹配算法类型。
            ////Sqdiff = 0 平方差匹配，最好的匹配为 0。
            ////SqdiffNormed = 1 归一化平方差匹配，最好效果为 0。
            ////Ccorr = 2 相关匹配法，数值越大效果越好。
            ////CcorrNormed = 3 归一化相关匹配法，数值越大效果越好。
            ////Ccoeff = 4 系数匹配法，数值越大效果越好。
            ////CcoeffNormed = 5 归一化系数匹配法，数值越大效果越好。
            #endregion
            CvInvoke.MatchTemplate(src, tempImg, matchImg, TemplateMatchingType.CcoeffNormed);
            #region 归一化函数参数说明
            ////IInputArray src：输入数据。
            ////IOutputArray dst：进行归一化后输出数据。
            ////double alpha = 1; 归一化后的最大值，默认为 1。
            ////double beta = 0：归一化后的最小值，默认为 0。
            #endregion
            CvInvoke.Normalize(matchImg, matchImg, 0, 1, NormType.MinMax, matchImg.Depth); //归一化
            double minValue = 0.0, maxValue = 0.0;
            Point minLoc = new Point();
            Point maxLoc = new Point();
            #region 极值函数参数说明
            ////IInputArray arr：输入数组。
            ////ref double minVal：输出数组中的最小值。
            ////ref double maxVal; 输出数组中的最大值。
            ////ref Point minLoc：输出最小值的坐标。
            ////ref Point maxLoc; 输出最大值的坐标。
            ////IInputArray mask = null：蒙版。
            #endregion
            CvInvoke.MinMaxLoc(matchImg, ref minValue, ref maxValue, ref minLoc, ref maxLoc);

            StringBuilder tb_result = new StringBuilder();
            tb_result.Append("min=" + minValue + ",max=" + maxValue);
            tb_result.Append(Environment.NewLine);
            tb_result.Append("最小值坐标：\n" + minLoc.ToString());
            tb_result.Append(Environment.NewLine);
            tb_result.Append("最大值坐标：\n" + maxLoc.ToString());
            SetText(tb_result.ToString());
            //Console.WriteLine(tb_result);
            CvInvoke.Rectangle(src, new Rectangle(maxLoc, tempImg.Size), new MCvScalar(0, 0, 255), 3);//绘制矩形，匹配得到的效果。

            CvInvoke.Imshow("result", src);
            CvInvoke.WaitKey(0);
        }

        private void button12_Click(object sender, EventArgs e)
        {

            try
            {
                ///图片颜色识别
                Mat src = CvInvoke.Imread("C:\\Users\\hudingwen\\Desktop\\副本\\20220824163014.jpg", LoadImageType.Color);


                ///手掌肤色提取 

                double h_min = 0, s_min = 70, v_min = 70;       //手的HSV颜色分布
                double h_max = 15, s_max = 255, v_max = 255;

                ScalarArray hsv_min = new ScalarArray(new MCvScalar(h_min, s_min, v_min));
                ScalarArray hsv_max = new ScalarArray(new MCvScalar(h_max, s_max, v_max));

                Mat hsvimg = new Mat();
                Mat mask = new Mat();

                CvInvoke.CvtColor(src, hsvimg, ColorConversion.Bgr2Hsv);
                CvInvoke.InRange(hsvimg, hsv_min, hsv_max, mask);
                CvInvoke.MedianBlur(mask, mask, 5);

                VectorOfVectorOfPoint contours = new VectorOfVectorOfPoint();
                VectorOfRect hierarchy = new VectorOfRect();
                //发现轮廓
                CvInvoke.FindContours(mask, contours, hierarchy, RetrType.External, ChainApproxMethod.ChainApproxNone);

                for (int i = 0; i < contours.Size; i++)
                {
                    Rectangle rect = CvInvoke.BoundingRectangle(contours[i]);
                    if (rect.Width < 10 || rect.Height < 10)
                        continue;
                    CvInvoke.Rectangle(src, rect, new MCvScalar(255, 0, 0));
                    CvInvoke.PutText(src, "hand", new Point(rect.X, rect.Y - 5), FontFace.HersheyComplexSmall, 1.2, new MCvScalar(0, 0, 255));
                }
                CvInvoke.Imshow("hsv_track", src);

                CvInvoke.WaitKey(0);
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            try
            {
                SetText("1231123");
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        private void comTask_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                jobName.Text = comTask.SelectedText;
                isStaring = false;
                //获取配置
                ReadConfig();
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            try
            {
                if(string.IsNullOrWhiteSpace(jobName.Text)) throw new Exception("任务名称不能为空");
               

                var dic = Path.Combine(Application.StartupPath, "jobs");
                if (!Directory.Exists(dic)) Directory.CreateDirectory(dic);

                var filePath = Path.Combine(dic, jobName.Text + ".config");
                if (File.Exists(filePath)) throw new Exception("任务重复");

                var config = Newtonsoft.Json.JsonConvert.SerializeObject(jobs, Formatting.Indented);
                File.WriteAllText(filePath, config);
                ReadTask();
                SetText("添加成功");
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        private void btnRm_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists(comTask.SelectedValue.ToString()))
                {
                    File.Delete(comTask.SelectedValue.ToString());
                }

                ReadTask();
                SetText("删除成功");
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        private void btRefresh_Click(object sender, EventArgs e)
        {
            try
            {
                ReadTask();
                SetText("刷新成功");
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        private void btnReCount_Click(object sender, EventArgs e)
        {
            try
            {
                var row = gridJobs.SelectedRows;
                if (row.Count > 0)
                {
                    List<job> ls = new List<job>();
                    foreach (DataGridViewRow item in row)
                    {
                        var tjob = (job)item.DataBoundItem;
                        ls.Add(tjob);
                    }
                    foreach (var item in ls)
                    {
                        item.countLess = 0;
                        item.count = Convert.ToInt32(numCount.Value);
                    }

                    RefreshData();
                    SaveConfig();
                    SetText($"次数设置成功");
                }
                else
                {
                    foreach (var item in jobs)
                    {
                        item.countLess = 0;
                        item.count = Convert.ToInt32(numCount.Value);
                    }
                }
                
                RefreshData();
                SaveConfig();

            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }
    }
}
