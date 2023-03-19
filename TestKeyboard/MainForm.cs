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
using Gma.QrCodeNet.Encoding.Windows.Render;
using Gma.QrCodeNet.Encoding;
using System.Drawing.Imaging;
using System.Configuration;
using DevExpress.XtraGrid.Columns;
using System.Media;
using DevExpress.XtraEditors.Controls;

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
        private Foreground mFocus = new Foreground();//聚焦
        private PressClick mClick = new PressClick();//鼠标点击
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
        public string pushApi;
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
            try
            {
                viewJobs.OptionsBehavior.ReadOnly = true;
                viewJobs.OptionsView.ShowGroupPanel = false;
                viewJobs.OptionsView.ColumnAutoWidth = false;
                viewJobs.OptionsBehavior.AutoPopulateColumns = false;
                viewJobs.OptionsSelection.MultiSelect = true;
                viewJobs.Columns.Clear();
                viewJobs.Columns.Add(new GridColumn { FieldName = "type", Caption = "任务", Width = 200, Visible = true });
                viewJobs.Columns.Add(new GridColumn { FieldName = "remark", Caption = "备注", Width = 200, Visible = true });
                viewJobs.Columns.Add(new GridColumn { FieldName = "time", Caption = "间隔(ms)", Width = 200, Visible = true });
                viewJobs.Columns.Add(new GridColumn { FieldName = "less", Caption = "下次(ms)", Width = 200, Visible = true });
                viewJobs.Columns.Add(new GridColumn { FieldName = "content", Caption = "内容", Width = 200, Visible = true });
                viewJobs.Columns.Add(new GridColumn { FieldName = "x", Caption = "X坐标", Width = 200, Visible = true });
                viewJobs.Columns.Add(new GridColumn { FieldName = "y", Caption = "Y坐标", Width = 200, Visible = true });
                viewJobs.Columns.Add(new GridColumn { FieldName = "duration", Caption = "持续时间(ms)", Width = 200, Visible = true });
                viewJobs.Columns.Add(new GridColumn { FieldName = "delay", Caption = "后置延迟(ms)", Width = 200, Visible = true });
                viewJobs.Columns.Add(new GridColumn { FieldName = "count", Caption = "运行次数", Width = 200, Visible = true });
                viewJobs.Columns.Add(new GridColumn { FieldName = "countLess", Caption = "已运行次数", Width = 200, Visible = true });
                viewJobs.Columns.Add(new GridColumn { FieldName = "isAsync", Caption = "并行任务", Width = 200, Visible = true });

                comMusic.SelectedIndex = 0;
                comMusic.Properties.TextEditStyle = TextEditStyles.DisableTextEditor;
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
            pushApi = ConfigurationManager.AppSettings["pushApi"];
            //获得微信绑定id
            CheckID();
            //开启日志显示
            Task.Run(ShowLog);
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

            //绑定数据源  
            gridJobs.DataSource = jobs;


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
        private void gridJobs_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            //Rectangle rectangle = new Rectangle(e.RowBounds.Location.X,
            //    e.RowBounds.Location.Y,
            //    gridJobs.RowHeadersWidth - 4,
            //    e.RowBounds.Height);
            //TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(),
            //    gridJobs.RowHeadersDefaultCellStyle.Font,
            //    rectangle,
            //    gridJobs.RowHeadersDefaultCellStyle.ForeColor,
            //    TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }
        ConcurrentBag<string> list = new ConcurrentBag<string>();
        public void SetText(string logStr, bool isClean = false)
        {
            try
            {
                if (logStr == null || string.IsNullOrWhiteSpace(logStr)) return;

                queue.Enqueue($"{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}-{logStr}\r\n");
            }
            catch (Exception)
            {

            }

        }
        public void ShowLog()
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
        public void WriteTask()
        {
            try
            {
                while (isStaring)
                {
                    Thread.Sleep(10000);
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

                for (int i = 0; i < jobs.Count; i++)
                {
                    Job item = jobs[i];
                    if (item.count > 0 && item.countLess >= item.count)
                        continue;
                    if (!(item.less <= 0 || item.less == item.time))
                        continue;
                    item.isCanStar = true;
                }
                //时间统计
                Task.Run(CaclTime);
                //刷新显示
                Task.Run(RefrData);
                //任务计算
                Task.Run(CalcData); 
                //定时回写
                Task.Run(WriteTask);
                //定时任务
                Task.Run(RunJob);
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }
        HttpClient httpClient = new HttpClient();

        public void RunJob()
        {
            SetText("任务开启");
            try
            {

                //循环任务
                while (isStaring)
                {

                    for (int i = 0; i < jobs.Count; i++)
                    {
                        if (!isStaring)
                            return;
                        Job item = jobs[i];
                        if (item.isCanStar)
                        {
                            item.isCanStar = false;
                            if (item.isAsync)
                                Task.Run(() => { StarJob(item); });
                            else
                                StarJob(item);
                        }
                        Thread.Sleep(10);
                    }
                    Thread.Sleep(10);
                    isFirst = false;
                }
            }
            catch (Exception ex)
            {
                SetText($"任务异常退出:{ex.Message}");
                isStaring = false;
            }
        }

        private void StarJob(Job item)
        {
            try
            {
                item.typeName = item.type.ToString();
                string msg = "";
                if (item.type == JobType.按键任务)
                {

                    var key = (KeyBoard)item.content;
                    if (item.duration > 0)
                        key_down((int)key);
                    else
                        click((int)key);
                    //if (item.children != null && item.children.Count > 0)
                    //{
                    //    foreach (var child in item.children)
                    //    {
                    //        StarJob(child);
                    //    }
                    //}
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
                        SetText("窗口聚焦失败");
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
                else if (item.type == JobType.截图检测)
                {
                    var screen = ScreenCat.GetScreen(item.content.ToString());
                    var savePath = Path.Combine(Application.StartupPath, "pics", $"{DateTime.Now.ToString("yyyMMddHHmmss")}.jpg");
                    ScreenCat.SaveImageWithQuality(savePath, screen, 70);
                    picCut.Image = screen;
                    SetText($"截图成功:{savePath}");
                    Task.Run(() => { CheckPicture(savePath, CheckType.一般测谎); });
                    Task.Run(() => { CheckPicture(savePath, CheckType.解轮检测); });
                    Task.Run(() => { CheckPicture(savePath, CheckType.蘑菇测谎); });
                }
                //后置延迟
                if (item.delay > 0)
                    Thread.Sleep((int)item.delay);
                //执行次数记录
                item.countLess += 1;

            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        private void CheckPicture(string savePath, CheckType checkType)
        {
            string msg_success = String.Empty;
            string msg_error = String.Empty;
            using (Process p = new Process())
            {
                p.StartInfo.UseShellExecute = false;  // 如果使用StandardOutput接收，这项必须为false（来自官方文档）
                p.StartInfo.CreateNoWindow = true;  // 是否创建窗口，true为不创建
                p.StartInfo.RedirectStandardOutput = true;  // 使用StandardOutput接收，一定要重定向标准输出，否则会报InvalidOperationException异常 
                p.StartInfo.RedirectStandardInput = true;
                p.StartInfo.RedirectStandardError = true;

                //p.StartInfo.StandardOutputEncoding = Encoding.UTF8;

                //p.StartInfo.StandardErrorEncoding = Encoding.UTF8;



                var myScript = Path.Combine(Application.StartupPath, "script", "043-find_lun.exe");


                var mySample = "";
                if (checkType == CheckType.解轮检测)
                    mySample = Path.Combine(Application.StartupPath, "sample", "lun.png");
                else if (checkType == CheckType.蘑菇测谎)
                    mySample = Path.Combine(Application.StartupPath, "sample", "mogu.png");
                else if (checkType == CheckType.一般测谎)
                {
                    myScript = Path.Combine(Application.StartupPath, "script", "045-blob_4.exe");
                    mySample = savePath;
                }
                else
                    return;
                p.StartInfo.FileName = myScript; // 设置要执行的Python脚本名称
                var myResource = savePath;
                p.StartInfo.Arguments = $"\"{mySample}\" \"{myResource}\"";  // 设置参数
                p.Start();  // 启动进程
                msg_success = p.StandardOutput.ReadToEnd();  // 接收信息 
                msg_error = p.StandardError.ReadToEnd(); //接收错误信息
                if (string.IsNullOrWhiteSpace(msg_success))
                {
                    //没有正常返回
                    SetText($"检测出错-{checkType}-{msg_error}");
                }
                else
                {
                    //正常返回
                    SetText($"{checkType}-{msg_success}");
                    MessageModel resMsg = new MessageModel();
                    try
                    {
                        resMsg = Newtonsoft.Json.JsonConvert.DeserializeObject<MessageModel>(msg_success);
                    }
                    catch (Exception)
                    {

                    }
                    if (resMsg.success)
                    {

                        //成功识别

                        //微信推送
                        if (!string.IsNullOrEmpty(wechatid))
                        {
                            var url = $"{pushApi}/api/WeChat/PushCardMsgGet?info.id=test&info.companyCode=test&info.userID={wechatid}&cardMsg.template_id=_eLuhnhPQMqVagS3SUyIlbYfi0gltTd4R0Jiq69gJA4&cardMsg.first={checkType.ToString()}-{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}\r\n截图地址:{Path.GetFileName(myResource)}";
                            try
                            {
                                string res = httpClient.GetStringAsync(url).Result;
                                WechatModel data = JsonConvert.DeserializeObject<WechatModel>(res);
                                if (data.success)
                                {
                                    SetText("微信消息推送成");
                                }
                            }
                            catch (Exception ex)
                            {
                                SetText($"推送错误{ex.Message}");
                                SetText($"推送错误{ex.StackTrace}");
                            }
                        }
                        //报警
                        if (checkMusic.Checked)
                        {
                            try
                            {
                                sp.SoundLocation = Path.Combine(Application.StartupPath, "music", $"{comMusic.Text}.wav");
                                sp.PlayLooping();
                            }
                            catch (Exception ex)
                            {
                                SetText(ex.Message);
                            }
                        }
                        
                    }
                }
                p.WaitForExit();
                p.Close();
                p.Dispose();
            }
        }

        public void RefrData()
        {
            try
            {
                SetText("开启刷新");
                while (isStaring)
                {
                    Thread.Sleep(1000);
                    RefreshData();
                }
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }


        public void CalcData()
        {
            //刷新
            try
            {
                SetText("开启计算");
                while (isStaring)
                {
                    Thread.Sleep(10);
                    for (int i = 0; i < jobs.Count; i++)
                    {
                        Job item = jobs[i];
                        item.less = item.less - 15;
                        if (item.less < 0) item.less = item.time;
                        if (!isStaring)
                            return;
                        if (item.count > 0 && item.countLess >= item.count)
                            continue;
                        if (item.less <= 0 || item.less == item.time)
                            item.isCanStar = true;
                    }
                }
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }

        } 
        public void CaclTime()
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

                Job job = new Job();
                jobs.Add(job);
                SetTime(job);
                SetTask(job);

                RefreshData();
                SetText("添加成功");
                SaveConfig();
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        private void SetTask(Job job)
        {
            var jtype = (KeyValuePair<int, JobType>)comJob.SelectedItem;
            if (jtype.Key == JobType.按键任务.GetHashCode())
            {
                job.content = (KeyBoard)this.comBoard.SelectedValue;
            }
            else if (jtype.Key == JobType.点击任务.GetHashCode())
            {
                job.content = (EnumMouse)this.comMouse.SelectedValue;
                job.x = (int)posX.Value;
                job.y = (int)posY.Value;
            }
            else if (jtype.Key == JobType.聚焦窗体.GetHashCode())
            {
                job.content = this.windowName.Text;
            }
            else if (jtype.Key == JobType.移动窗体.GetHashCode())
            {
                job.content = this.windowName.Text;
                job.x = Convert.ToInt32(posX.Value);
                job.y = Convert.ToInt32(posY.Value);
            }
            else if (jtype.Key == JobType.截图检测.GetHashCode())
            {
                job.content = this.windowName.Text;
            }
            job.isAsync = checkAysnc.Checked;
            job.type = (JobType)this.comJob.SelectedValue;
            job.contentName = job.content.ToString();
            job.typeName = job.type.ToString();
            job.count = Convert.ToInt32(numCount.Value);
            job.remark = txtRemark.Text;
        }

        private void SetTime(Job job)
        {
            job.time = (int)(Convert.ToDecimal(numTime.Text) * 1000);
            job.less = (int)(Convert.ToDecimal(numTime.Text) * 1000);
            job.delay = (int)(Convert.ToDecimal(numDelay.Text) * 1000);
            job.duration = (int)(Convert.ToDecimal(textDuration.Text) * 1000);
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
            //if (item.children != null && item.children.Count > 0)
            //{
            //    foreach (var child in item.children)
            //    {
            //        ConvertJobType(child);
            //    }
            //}
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("确定要删除?", "提示", MessageBoxButtons.OKCancel) != DialogResult.OK) return;
                var row = viewJobs.GetSelectedRows();
                if (row.Length > 0)
                {
                    List<Job> ls = new List<Job>();
                    foreach (int idx in row)
                    {
                        var tjob = (Job)viewJobs.GetRow(idx);
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

            //this.gridJobs.BeginInvoke(new Action(() =>
            //{
            //    source.ResetBindings(false);
            //}));
            //gridJobs.DataSource = null;
            //source = new BindingSource();
            //source.DataSource = jobs;
            //gridJobs.DataSource = jobs;  

            //gridJobs.RefreshDataSource();


            this.BeginInvoke(new Action(() =>
            {
                viewJobs.RefreshData();
                viewJobs.BestFitColumns();
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
            Task.Run(TestScreen);
        }

        Process p;
        public void TestScreen()
        {
            try
            {
                Thread.Sleep(1000);
                //var screen = ScreenCat.GetScreen(checkScreen.Checked, (int)posX.Value, (int)posY.Value, (int)numWidth.Value, (int)numHeight.Value);
                var msg = ""; 
                mFocus.FocusWindow(windowName.Text, ref msg);
                var screen = ScreenCat.GetScreen(windowName.Text);
                var savePath = Path.Combine(Application.StartupPath, "pics", $"{DateTime.Now.ToString("yyyMMddHHmmss")}.jpg");
                ScreenCat.SaveImageWithQuality(savePath, screen, 100);
                picCut.Image = screen;
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
                sp.Stop();
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }
        System.Media.SoundPlayer sp = new SoundPlayer();
        private void btntest_stop_Click(object sender, EventArgs e)
        {
            try
            {
                sp.SoundLocation = Path.Combine(Application.StartupPath, "music", $"{comMusic.Text}.wav");
                sp.PlayLooping();
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
                if (string.IsNullOrWhiteSpace(jobName.Text)) throw new Exception("任务组名称不能为空");


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
                var row = viewJobs.GetSelectedRows();
                if (row.Length > 0)
                {
                    List<Job> ls = new List<Job>();
                    foreach (int idx in row)
                    {
                        var tjob = (Job)viewJobs.GetRow(idx);
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
                var row = viewJobs.GetSelectedRows();
                if (row.Length > 0)
                {
                    List<Job> ls = new List<Job>();
                    foreach (int item in row)
                    {
                        var tjob = (Job)viewJobs.GetRow (item);
                        ls.Add(tjob);
                    }
                    foreach (var item in ls)
                    {
                        SetTime(item);
                    }
                    SetText($"时间设置成功");
                }
                else
                {
                    SetText($"请选择要设置时间的任务");
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
                var rows = viewJobs.GetSelectedRows();
                if (!(rows != null && rows.Length == 1))
                {
                    SetText("请选择一个数据");
                    return;
                }
                var curRow = (Job)viewJobs.GetRow(rows[0]);
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
                var rows = viewJobs.GetSelectedRows();
                if (!(rows != null && rows.Length == 1))
                {
                    SetText("请选择一个数据");
                    return;
                }
                var curRow = (Job)viewJobs.GetRow(rows[0]);
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
                    //txtSample.Text = openFileDialog.FileName;
                }
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        public string wechatid { get; set; }

        private void btn_wechat_Click(object sender, EventArgs e)
        {
            try
            {
                System.Guid guid = new Guid();
                guid = Guid.NewGuid();
                string id = guid.ToString("N");
                var url = $"{pushApi}/api/WeChat/GetQRBind?id=test&companyCode=test&userID={id}";

                var res = httpClient.GetStringAsync(url).Result;

                WechatModel data =JsonConvert.DeserializeObject<WechatModel>(res);
                if (data.success)
                {
                    var dic = Path.Combine(Application.StartupPath, "config");
                    if (!Directory.Exists(dic)) Directory.CreateDirectory(dic);
                    var config = Path.Combine(dic, "id.config");
                    File.WriteAllText(config, id);

                    QrEncoder qrEncoder = new QrEncoder(ErrorCorrectionLevel.H);
                    QrCode qrCode = new QrCode();
                    qrEncoder.TryEncode(data.response.usersData.url, out qrCode);

                    GraphicsRenderer renderer = new GraphicsRenderer(new FixedModuleSize(5, QuietZoneModules.Two), Brushes.Black, Brushes.White);

                    using (MemoryStream ms = new MemoryStream())
                    {
                        renderer.WriteToStream(qrCode.Matrix, ImageFormat.Png, ms);
                        Image img = Image.FromStream(ms);
                        picWeChat.Image = img;
                    }
                    btn_wechat.Enabled = false;
                    wechatid = id;
                    Task.Run(CheckIsBind);
                }
                SetText("获取绑定二维码成功");

            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        public void CheckIsBind()
        {
            try
            {
                int i = 0;
                while (i < 10)
                {
                    Thread.Sleep(10000);
                    var url = $"{pushApi}/api/WeChat/GetBindUserInfo?id=test&companyCode=test&userID={wechatid}";
                    var res = httpClient.GetStringAsync(url).Result;
                    WechatModel data = JsonConvert.DeserializeObject<WechatModel>(res);
                    if (data.success)
                    {
                        SetText("微信绑定成功");
                        break;
                    }
                    i++;
                }
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }
        public void CheckID()
        {
            try
            {
                var dic = Path.Combine(Application.StartupPath, "config");
                if (!Directory.Exists(dic)) Directory.CreateDirectory(dic);
                var config = Path.Combine(dic, "id.config");
                if(File.Exists(config))
                {
                    btn_wechat.Enabled = false;
                    wechatid = File.ReadAllText(config);
                }
                
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        private void btnSetTask_Click(object sender, EventArgs e)
        {
            try
            {
                var row = viewJobs.GetSelectedRows();
                if (row.Length == 1)
                {
                    var item = (Job)viewJobs.GetRow(row[0]);   
                    SetTask(item);
                    SetText($"任务设置成功");
                }
                else
                {
                    SetText($"请选择一个要重新设置的任务");
                }

                RefreshData();
                SaveConfig();

            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        private void btnStopMusic_Click(object sender, EventArgs e)
        {
            try
            {
                sp.Stop();
                SetText("停止播放");
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }
    }
}
