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
using System.Data;
using System.Net;
using System.Threading.Tasks;

namespace TestKeyboard
{

    public partial class MainForm : Form
    {

        [DllImport("librustdesk.dll")]
        public static extern bool click(int keycode);
        [DllImport("librustdesk.dll")]
        public static extern bool press(int keycode, int time);
        [DllImport("librustdesk.dll")]
        public static extern bool key_down(int keycode);
        [DllImport("librustdesk.dll")]
        public static extern bool key_up(int keycode);
        private Foreground mFocus;
        private PressClick mClick;
        static ConcurrentQueue<string> queue = new ConcurrentQueue<string>();
        /// <summary>
        /// 键盘钩子KeyDown事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void HookMain_OnKeyDown(object sender, KeyEventArgs e)
        {
            SetText("键盘输入" + e.KeyCode);
            if (e.KeyCode == Keys.F1)
            {
                RemoveEvent_key();
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
                    if (isStaring)
                    {

                        StopTask();
                    }
                    else
                    {
                        Start();
                    }

                    SaveConfig();
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
            this.FormClosing += MainForm_FormClosing;
            Control.CheckForIllegalCrossThreadCalls = false;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                StopTask();
                RemoveEvent();
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        public DateTime runTime = DateTime.MinValue;//运行时间 
        public bool isStaring = false;//开启标志
        public List<Job> jobs = new List<Job>();
        BindingSource source;
        //任务类型
        Dictionary<int, JobType> dict = new Dictionary<int, JobType>();
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

            //开启监听
            AddEvent();

            runTime = DateTime.Now;
            //任务类型
            foreach (JobType item in Enum.GetValues(typeof(JobType)))
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
        public void SetText(string logStr, bool isClean = false)
        {
            try
            {
                if (logStr == null || string.IsNullOrWhiteSpace(logStr)) return;
               
                queue.Enqueue($"{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}-{logStr}\r\n");
                //Action log = new Action(() => { });
                //log.BeginInvoke(ShowLog, logStr);
            }
            catch (Exception)
            {
               
            }

        }
        public void ShowLog(IAsyncResult ar)
        {
            try
            {
                while (true)
                {
                    if (queue.Count > 0)
                    {
                        string log;
                        var isok = queue.TryDequeue(out log);
                        if (isok && !string.IsNullOrEmpty(log))
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
            //开启任务
            Start();
        }
        public void Start()
        {
            try
            {
                SetText("任务开始", true);
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

                //循环任务
                while (isStaring)
                {
                    DateTime starTime = DateTime.Now;
                    var lessJobs = jobs.Where(t => t.count > 0 && t.countLess >= t.count).ToList();
                    if (lessJobs.Count == jobs.Count)
                    {
                        isStaring = false;
                        SetText("所有任务执行结束");
                        break;
                    }
                    for (int i = 0; i < jobs.Count; i++)
                    {
                        if (!isStaring)
                            break;
                        Job item = jobs[i];
                        if (item.count > 0 && item.countLess >= item.count)
                            continue;
                        StarJob(item);

                    }

                    var total = (int)(DateTime.Now - starTime).TotalMilliseconds;
                    foreach (var other in jobs)
                    {
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

        private void StarJob(Job item)
        {
            try
            {
                item.typeName = item.type.ToString();

                string msg = "";

                if (item.less <= 0 || item.less == item.time)
                {
                    if (item.type == JobType.按键任务)
                    {

                        var key = (KeyBoard)item.content;
                        if (item.duration > 0)
                            key_down((int)key);
                        else
                            click((int)key);
                        if (item.children != null && item.children.Count > 0)
                        {
                            foreach (var child in item.children)
                            {
                                StarJob(child);
                            }
                        }
                        if (item.duration > 0)
                        {
                            Thread.Sleep(item.duration);
                            key_up((int)key);
                        }
                        //取反
                        if (key == KeyBoard.LeftArrow || key == KeyBoard.RightArrow)
                        {

                            if (key == KeyBoard.LeftArrow)
                                item.content = KeyBoard.RightArrow;
                            else if (key == KeyBoard.RightArrow)
                                item.content = KeyBoard.LeftArrow;
                        }
                        SetText("按键" + item.content.ToString());
                    }
                    else if (item.type == JobType.点击任务)
                    {

                        mClick.Click(item.x, item.y, (EnumMouse)item.content);
                        SetText("点击" + item.content.ToString());
                    }
                    else if (item.type == JobType.聚焦窗体)
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
                    else if (item.type == JobType.移动窗体)
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
                    else if (item.type == JobType.截图窗体)
                    {
                        var screen = ScreenCat.GetScreen(item.content.ToString());
                        var savePath = Path.Combine(Application.StartupPath, "pics", $"{DateTime.Now.ToString("yyyMMddHHmmss")}.jpg");
                        ScreenCat.SaveImageWithQuality(savePath, screen, 100);
                        pictureBox1.Image = screen;
                        SetText($"截图成功:{savePath}");

                        Task.Run(() => { CheckPicture(savePath, CheckType.一般测谎); });
                        Task.Run(() => { CheckPicture(savePath, CheckType.解轮检测); });
                        Task.Run(() => { CheckPicture(savePath, CheckType.蘑菇测谎); });
                        Task.Run(() => { CheckPicture(savePath, CheckType.蘑菇测谎_2); });
                        Task.Run(() => { CheckPicture(savePath, CheckType.蘑菇测谎_3); });
                        Task.Run(() => { CheckPicture(savePath, CheckType.蘑菇测谎_4); });
                    }
                    //后置延迟
                    if (item.delay > 0)
                        Thread.Sleep((int)item.delay);
                    //执行次数记录
                    item.countLess += 1;
                }
                else
                {
                    //还没到执行的时候

                }
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        private  void CheckPicture(string savePath, CheckType checkType)
        {
           


            string msg_success = String.Empty;
            string msg_error = String.Empty;
            using (Process p = new Process())
            {
                p.StartInfo.UseShellExecute = false;  // 如果使用StandardOutput接收，这项必须为false（来自官方文档）
                p.StartInfo.CreateNoWindow = true;  // 是否创建窗口，true为不创建
                p.StartInfo.RedirectStandardOutput = true;  // 使用StandardOutput接收，一定要重定向标准输出，否则会报InvalidOperationException异常 
                //p.StartInfo.RedirectStandardInput = true;
                p.StartInfo.RedirectStandardError = true;
                p.StartInfo.FileName = "python"; // 设置要执行的Python脚本名称

                var myScript = "D:/hudingwen/github/my-OpenCV-Python/043-find_lun.py";
                var mySample = "";
                if (checkType == CheckType.解轮检测)
                    mySample = "D:/hudingwen/github/my-OpenCV-Python/pic/mxd/test/lun.png";
                else if (checkType == CheckType.蘑菇测谎)
                    mySample = "D:/hudingwen/github/my-OpenCV-Python/pic/mxd/test/mogu.png";
                else if (checkType == CheckType.蘑菇测谎_2)
                    mySample = "D:/hudingwen/github/my-OpenCV-Python/pic/mxd/test/mogu2.png";
                else if (checkType == CheckType.蘑菇测谎_3)
                    mySample = "D:/hudingwen/github/my-OpenCV-Python/pic/mxd/test/mogu3.png";
                else if (checkType == CheckType.蘑菇测谎_4)
                    mySample = "D:/hudingwen/github/my-OpenCV-Python/pic/mxd/test/mogu4.png";
                else if (checkType == CheckType.一般测谎)
                {
                    myScript = "D:/hudingwen/github/my-OpenCV-Python/045-blob_4.py";
                    mySample = savePath;
                }
                else
                    return;
                //
                // var myResource = "D:/hudingwen/github/my-OpenCV-Python/pic/mxd/test/lun1.jpg";
                var myResource = savePath;
                p.StartInfo.Arguments = $"{myScript} {mySample} {myResource}";  // 设置参数
                p.Start();  // 启动进程
                msg_success = p.StandardOutput.ReadToEnd();  // 接收信息 
                msg_error = p.StandardError.ReadToEnd(); //接收错误信息
                if (string.IsNullOrWhiteSpace(msg_success))
                {
                    //没有正常返回
                    SetText(msg_error);
                }
                else
                {
                    //正常返回
                    SetText(msg_success);
                    MessageModel resMsg  = new MessageModel();
                    try
                    {
                        resMsg = Newtonsoft.Json.JsonConvert.DeserializeObject<MessageModel>(msg_success);
                    }
                    catch (Exception)
                    { 
                        
                    }
                    if (resMsg.success)
                    {
                        //SetText($"脚本地址:{myScript}");
                        //SetText($"样本地址:{mySample}");
                        //SetText($"截图地址:{myResource}");
                        //成功识别
                        var url = $"";

                        string ret = string.Empty;
                        try
                        {
                            //ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls13;

                            //ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                            HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create(url);
                            webReq.Method = "GET";
                            webReq.ContentType = "application/json";
                            //webReq.Headers.Add("Authorization", "bearer ********");
                            //Stream postData = webReq.GetRequestStream();
                            //postData.Close();
                            HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
                            StreamReader sr = new StreamReader(webResp.GetResponseStream(), Encoding.UTF8);
                            ret = sr.ReadToEnd();
                            SetText($"微信推送:{ret}");
                        }
                        catch (Exception ex)
                        {
                            SetText($"推送错误{ex.Message}");
                            SetText($"推送错误{ex.StackTrace}");
                        }
                    }
                }
                p.WaitForExit();
                p.Close();
                p.Dispose();
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
                StopTask();
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }

        }

        public void StopTask()
        {
            isStaring = false;
            SaveConfig();
            SetText("任务停止");
        }


        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                Foreground win = new Foreground();
                string msg = "";
                win.FocusWindow(windowName.Text, ref msg);
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
                action.BeginInvoke(TestClick, null);

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
            SetText("点击x:" + (int)posX.Value + "|y:" + (int)posY.Value);
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
            SetText("开启鼠标监听");
            m_HookMain.OnMouseActivity += mouseEvent;
            m_HookMain.InstallHook();

        }
        public void RemoveEvent()
        {
            m_HookMain.OnKeyDown -= keyEvent;
            m_HookMain.OnMouseUpdate -= mouseUpdateEvent;
            m_HookMain.OnMouseActivity -= mouseEvent;
            m_HookMain.UnInstallHook();
            SetText("关闭鼠标监听");
        }

        public void AddEvent_Key()
        {
            keyEvent = new KeyEventHandler(HookMain_OnKeyDown);
            m_HookMain.OnKeyDown += keyEvent;
            SetText("开启键盘监听");
            m_HookMain.InstallHook_key();

        }
        public void RemoveEvent_key()
        {
            m_HookMain.OnKeyDown -= keyEvent;
            m_HookMain.UnInstallHook_key();
            isStaring = false;
            SetText("关闭键盘监听");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                isStaring = true;
                AddEvent_Key();
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
                var jtype = (KeyValuePair<int, JobType>)comJob.SelectedItem;
                Job job = new Job();
                if (jtype.Key == JobType.按键任务.GetHashCode())
                {
                    job = new Job
                    {
                        content = (KeyBoard)this.comBoard.SelectedValue,
                    };
                }
                else if (jtype.Key == JobType.点击任务.GetHashCode())
                {
                    job = new Job
                    {
                        content = (EnumMouse)this.comMouse.SelectedValue,
                        x = (int)posX.Value,
                        y = (int)posY.Value
                    };
                }
                else if (jtype.Key == JobType.聚焦窗体.GetHashCode())
                {
                    job = new Job
                    {
                        content = this.windowName.Text,
                    };
                }
                else if (jtype.Key == JobType.移动窗体.GetHashCode())
                {
                    job = new Job
                    {
                        content = this.windowName.Text,
                        x = Convert.ToInt32(posX.Value),
                        y = Convert.ToInt32(posY.Value),

                    };

                }
                else if (jtype.Key == JobType.截图窗体.GetHashCode())
                {
                    job = new Job
                    {
                        content = this.windowName.Text,
                        pathSample = this.txtSample.Text,
                    };
                }
                jobs.Add(job);
                job.type = (JobType)this.comJob.SelectedValue;
                job.time = (int)(Convert.ToDecimal(numTime.Text) * 1000);
                job.less = (int)(Convert.ToDecimal(numTime.Text) * 1000);
                job.delay = (int)(Convert.ToDecimal(numDelay.Text) * 1000);
                job.duration = (int)(Convert.ToDecimal(textDuration.Text) * 1000);
                job.contentName = job.content.ToString();
                job.typeName = job.type.ToString();
                job.count = Convert.ToInt32(numCount.Value);
                job.remark = txtRemark.Text;

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
                List<Job> ls = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Job>>(config);
                if (ls != null && ls.Count > 0)
                {
                    foreach (var item in ls)
                    {
                        ConvertJobType(item);
                    }
                    jobs.AddRange(ls);
                }
            }
            RefreshData();

        }

        private static void ConvertJobType(Job item)
        {
            if (item.type == JobType.按键任务 && item.content is long)
                item.content = (KeyBoard)Enum.ToObject(typeof(KeyBoard), Convert.ToInt32(item.content));
            if (item.type == JobType.点击任务 && item.content is long)
                item.content = (EnumMouse)Enum.ToObject(typeof(EnumMouse), Convert.ToInt32(item.content));
            if (item.children != null && item.children.Count > 0)
            {
                foreach (var child in item.children)
                {
                    ConvertJobType(child);
                }
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("确定要删除?", "提示", MessageBoxButtons.OKCancel) != DialogResult.OK) return;
                var row = gridJobs.SelectedRows;
                if (row.Count > 0)
                {
                    List<Job> ls = new List<Job>();
                    foreach (DataGridViewRow item in row)
                    {
                        var tjob = (Job)item.DataBoundItem;
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
                Thread.Sleep(1000);
                //var screen = ScreenCat.GetScreen(checkScreen.Checked, (int)posX.Value, (int)posY.Value, (int)numWidth.Value, (int)numHeight.Value);
                var msg = "";
                mFocus = new Foreground();
                mFocus.FocusWindow(windowName.Text, ref msg);
                var screen = ScreenCat.GetScreen(windowName.Text);
                var savePath = Path.Combine(Application.StartupPath, "pics", $"{DateTime.Now.ToString("yyyMMddHHmmss")}.jpg");
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
            try
            {

                var myScript = "D:/hudingwen/github/my-OpenCV-Python/043-find_lun.py";
                // 创建一个ProcessStartInfo对象
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    RedirectStandardInput = true,
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                };

                // 创建一个Process对象，并启动进程
                Process cmdProcess = new Process
                {
                    StartInfo = startInfo
                };
                cmdProcess.Start();

                // 获取标准输入和输出流
                var cmdIn = cmdProcess.StandardInput;
                var cmdOut = cmdProcess.StandardOutput;
                 
                cmdProcess.StandardInput.WriteLine("ping"); // 向Python发送输入
                cmdProcess.StandardInput.Flush();

                string output = cmdProcess.StandardOutput.ReadToEnd(); // 从Python获取输出
                Console.WriteLine(output);


            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
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
                if (string.IsNullOrWhiteSpace(jobName.Text)) throw new Exception("任务名称不能为空");


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
                if (MessageBox.Show("确定要删除?", "提示", MessageBoxButtons.OKCancel) != DialogResult.OK)
                    return;
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
                    List<Job> ls = new List<Job>();
                    foreach (DataGridViewRow item in row)
                    {
                        var tjob = (Job)item.DataBoundItem;
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
                    List<Job> ls = new List<Job>();
                    foreach (DataGridViewRow item in row)
                    {
                        var tjob = (Job)item.DataBoundItem;
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

        private void btnUp_Click(object sender, EventArgs e)
        {
            try
            {
                var rows = gridJobs.SelectedRows;
                if (!(rows != null && rows.Count == 1))
                {
                    SetText("请选择一个数据");
                    return;
                }
                var curRow = (Job)rows[0].DataBoundItem;
                int idx = jobs.FindIndex(t => t == curRow);
                MoveUp<Job>(jobs, idx);
                RefreshData();
                SaveConfig(); 

                
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        private void btnDown_Click(object sender, EventArgs e)
        {
            try
            {
                var rows = gridJobs.SelectedRows;
                if (!(rows != null && rows.Count == 1))
                {
                    SetText("请选择一个数据");
                    return;
                }
                var curRow = (Job)rows[0].DataBoundItem;
                int idx = jobs.FindIndex(t => t == curRow);
                MoveDown<Job>(jobs, idx);
                RefreshData();
                SaveConfig(); 


            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        // 将列表中索引为index的元素向上移动
        public static void MoveUp<T>(List<T> list, int index)
        {
            if (index > 0 && index < list.Count)
            {
                T temp = list[index];
                list[index] = list[index - 1];
                list[index - 1] = temp;
            }
        }

        // 将列表中索引为index的元素向下移动
        public static void MoveDown<T>(List<T> list, int index)
        {
            if (index >= 0 && index < list.Count - 1)
            {
                T temp = list[index];
                list[index] = list[index + 1];
                list[index + 1] = temp;
            }
        }

        private void btnSelectSample_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "png|*.png|jpg|*.jpg";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtSample.Text = openFileDialog.FileName;
                }
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }
    }
}
