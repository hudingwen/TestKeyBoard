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
using System.Text;
using System.Collections.Concurrent;
using System.Linq;
using System.Diagnostics;
using System.Net.Http;
using System.Runtime.InteropServices;

namespace TestKeyboard
{

    public partial class MainForm : Form
    {

        [DllImport("librustdesk.dll")]
        public static extern bool click(int keycode, int time); 
        private Foreground mFocus;
        private PressClick mClick;
        static Queue<string> queue = new Queue<string>();
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

        public MainForm()
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
            //开启日志显示
            Action log = new Action(() => { });
            log.BeginInvoke(ShowLog, null);

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
            if (files.Length == 0)
            {
                var config = Path.Combine(dic, "缺省.config");
                File.WriteAllText(config, "");
                files = Directory.GetFiles(dic, "*.config");
            }
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
                if (logStr == null || string.IsNullOrWhiteSpace(logStr)) return;
                queue.Enqueue($"{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}-{logStr}\r\n");
                //Action log = new Action(() => { });
                //log.BeginInvoke(ShowLog, logStr);
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
                while (true)
                {
                    if(queue.Count>0)
                    {
                        var log = queue.Dequeue();
                        if (log != null)
                        {
                            richLogs.AppendText(log);
                            richLogs.ScrollToCaret();
                        }
                    }
                    Thread.Sleep(100);
                }
                //if (ar.AsyncState == null) return;
                //string logStr = ar.AsyncState.ToString();
                //if (list.Count > 20)
                //{
                //    for (int i = 0; i < 30; i++)
                //    {
                //        string temp;
                //        list.TryTake(out temp);
                //    }
                //}
                //list.Add(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:ffff") + "-" + logStr);
                //richLogs.Text = string.Join("\r\n", list.OrderByDescending(t => t));
                

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
        bool isFirst = false;
        private void Button2_Click(object sender, EventArgs e)
        {
            //开启监听
            AddEvent();
            //开启任务
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
               
                mFocus = new Foreground();//聚焦
                mClick = new PressClick();//点击

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
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }
        HttpClient httpClient = new HttpClient();

        public async void RunJob(IAsyncResult ar)
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
                                try
                                {
                                    
                                    var key = (KeyBoard)item.content;
                                    click((int)key,Convert.ToInt32(item.duration));
                                    //取反
                                    if (key == KeyBoard.LeftArrow || key == KeyBoard.RightArrow)
                                    {
                                        
                                        if (key == KeyBoard.LeftArrow)
                                            item.content = KeyBoard.RightArrow;
                                        else if (key == KeyBoard.RightArrow)
                                            item.content = KeyBoard.LeftArrow;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    SetText(ex.Message);
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
                            if (isDely)
                                Thread.Sleep((int)item.delay);

                            //执行次数记录
                            item.countLess += 1;
                        }
                        else
                        {
                           //还没到执行的时候

                        }

                       
                    }
                    stopwatch.Stop();
                    foreach (var other in jobs)
                    {
                        var total = (int)stopwatch.ElapsedMilliseconds;
                        other.less -= total;
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
                        content = (KeyBoard)this.comBoard.SelectedValue,
                    };
                }
                else if (jtype.Key == jobType.点击任务.GetHashCode())
                {
                    job = new job
                    {
                        content = (EnumMouse)this.comMouse.SelectedValue,
                        x = (int)posX.Value,
                        y = (int)posY.Value
                    };
                }
                else if (jtype.Key == jobType.聚焦窗体.GetHashCode())
                {
                    job = new job
                    {
                        content = this.windowName.Text,
                    };
                }
                else if (jtype.Key == jobType.移动窗体.GetHashCode())
                {
                    job = new job
                    {
                        content = this.windowName.Text,
                        x = Convert.ToInt32(posX.Value),
                        y = Convert.ToInt32(posY.Value),
                        
                    };
                   
                }
                jobs.Add(job);
                job.type = (jobType)this.comJob.SelectedValue;
                job.time = (int)(Convert.ToDecimal(numTime.Text) * 1000);
                job.less = (int)(Convert.ToDecimal(numTime.Text) * 1000);
                job.delay = (int)(Convert.ToDecimal(numDelay.Text) * 1000);
                job.duration = (int)(Convert.ToDecimal(textDuration.Text) * 1000);
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
                    item.contentName = item.content.ToString();
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

        Process p;
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
        private void btnTest_Click(object sender, EventArgs e)
        {
           
        }
        private void btntest_stop_Click(object sender, EventArgs e)
        {
            
        
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

        private void btnSetTime_Click(object sender, EventArgs e)
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
                        item.time = (int)(Convert.ToDecimal(numTime.Text) * 1000);
                        item.less = (int)(Convert.ToDecimal(numTime.Text) * 1000);
                        item.delay = (int)(Convert.ToDecimal(numDelay.Text) * 1000);
                        item.duration = (int)(Convert.ToDecimal(textDuration.Text) * 1000);
                    }

                    RefreshData();
                    SaveConfig();
                    SetText($"时间设置成功");
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
