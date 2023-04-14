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
using System.Collections.Concurrent; 
using System.Diagnostics;
using System.Net.Http;
using System.Runtime.InteropServices; 
using System.Threading.Tasks;
using Gma.QrCodeNet.Encoding.Windows.Render;
using Gma.QrCodeNet.Encoding; 
using System.Configuration;
using DevExpress.XtraGrid.Columns;
using System.Media;
using DevExpress.XtraEditors.Controls;  
using System.Collections;
using DevExpress.CodeParser;
using System.Runtime.InteropServices.ComTypes;

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
                viewJobs.Columns.Add(new GridColumn { FieldName = "width", Caption = "宽度", Width = 200, Visible = true });
                viewJobs.Columns.Add(new GridColumn { FieldName = "height", Caption = "高度", Width = 200, Visible = true });
                viewJobs.Columns.Add(new GridColumn { FieldName = "duration", Caption = "持续时间(ms)", Width = 200, Visible = true });
                viewJobs.Columns.Add(new GridColumn { FieldName = "delay", Caption = "后置延迟(ms)", Width = 200, Visible = true });
                viewJobs.Columns.Add(new GridColumn { FieldName = "count", Caption = "运行次数", Width = 200, Visible = true });
                viewJobs.Columns.Add(new GridColumn { FieldName = "countLess", Caption = "已运行次数", Width = 200, Visible = true });
                viewJobs.Columns.Add(new GridColumn { FieldName = "isAsync", Caption = "并行任务", Width = 200, Visible = true });
                viewJobs.Columns.Add(new GridColumn { FieldName = "enable", Caption = "启用", Width = 200, Visible = true });


                comMusic.SelectedIndex = 0;
                comMusic.Properties.TextEditStyle = TextEditStyles.DisableTextEditor;

                
                comTask.SelectedValueChanged += comTask_SelectedValueChanged;


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
            comJob.Properties.Items.Clear();
            foreach (JobType item in Enum.GetValues(typeof(JobType)))
            {
                comJob.Properties.Items.Add(item.ToString(),item,0);
            }
            //键盘类型 
            foreach (KeyBoard item in Enum.GetValues(typeof(KeyBoard)))
            {
                var name = item.ToString();
                if (name.Equals("空格键"))
                {

                }
                comBoard.Properties.Items.Add(item.ToString(), item, 0);
            }
            //鼠标类型 
            foreach (EnumMouse item in Enum.GetValues(typeof(EnumMouse)))
            {
                comMouse.Properties.Items.Add(item.ToString(), item, 0);
            } 
            //任务列表
            ReadTask();

            //绑定数据源  
            gridJobs.DataSource = jobs;


            //设置默认
            
            comJob.SelectedIndex = 0;
            comBoard.SelectedIndex  = 0;
            comMouse.SelectedIndex = 0; 




        }

        public void ReadTask(string taskName="")
        {
            dicjobs.Clear();
            comTask.Properties.Items.Clear();
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
            foreach(var item in dicjobs)
            {
                comTask.Properties.Items.Add(item.Key,item.Value,0);
            }


            try
            { 
                var configPath = Path.Combine(Application.StartupPath, "config", "settings.config");
                if (File.Exists(configPath))
                {
                    var settings = File.ReadAllText(configPath);
                    var setInfo = Newtonsoft.Json.JsonConvert.DeserializeObject<ConfigInfo>(settings);
                    
                    checkMusic.Checked = setInfo.isPlay;
                    comMusic.Text = setInfo.playName;
                    if (!string.IsNullOrEmpty(taskName))
                        setInfo.taskName = taskName;
                    var taskIdx = 0;
                    bool isEnter = false;
                    foreach (var item in dicjobs)
                    {
                        if (item.Key.Equals(setInfo.taskName))
                        {
                            comTask.SelectedIndex = taskIdx;
                            isEnter = true;
                            break;
                        }
                        taskIdx++;
                    }
                    if (!isEnter)
                        comTask.SelectedIndex = 0;
                }
                else
                {
                    var taskIdx = 0;
                    bool isEnter = false;
                    foreach (var item in dicjobs)
                    {
                        if (item.Key.Equals(taskName))
                        {
                            comTask.SelectedIndex = taskIdx;
                            isEnter = true;
                            break;
                        }
                        taskIdx++;
                    }
                    if (!isEnter)
                        comTask.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }

        } 
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
                            richLogs.BeginInvoke(new Action(() =>
                            { 
                                richLogs.AppendText(log);
                                richLogs.ScrollToCaret();
                            })); 
                        }
                    }
                    Thread.Sleep(100);
                } 

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
                    if (item.less == item.time)
                        item.isFirstRun = true;
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
                //定时删除截图文件
                Task.Run(DeleHistoryPics);
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
                        if ((item.less == 0 || item.isFirstRun) && (item.count == 0 || (item.count > 0 && item.count > item.countLess)) && item.enable)
                        {
                            item.isFirstRun = false;
                            if (item.isAsync)
                            {
                                Task.Run(() => { StarJob(item); });
                            }
                            else
                            {
                                StarJob(item);
                            }
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

                    if (key == KeyBoard.左右走)
                    {
                        if (item.flag)
                        {
                            if (item.duration > 0)
                                key_down((int)KeyBoard.LeftArrow);
                            else
                                click((int)key);
                            if (item.duration > 0)
                            {
                                Thread.Sleep(item.duration);
                                key_up((int)KeyBoard.LeftArrow);
                            }
                            Thread.Sleep(item.delay);
                            if (item.duration > 0)
                                key_down((int)KeyBoard.RightArrow);
                            else
                                click((int)key);
                            if (item.duration > 0)
                            {
                                Thread.Sleep(item.duration);
                                key_up((int)KeyBoard.RightArrow);
                            }
                        }
                        else
                        {
                            if (item.duration > 0)
                                key_down((int)KeyBoard.RightArrow);
                            else
                                click((int)key);
                            if (item.duration > 0)
                            {
                                Thread.Sleep(item.duration);
                                key_up((int)KeyBoard.RightArrow);
                            }
                            Thread.Sleep(item.delay);
                            if (item.duration > 0)
                                key_down((int)KeyBoard.LeftArrow);
                            else
                                click((int)key);
                            if (item.duration > 0)
                            {
                                Thread.Sleep(item.duration);
                                key_up((int)KeyBoard.LeftArrow);
                            }
                        }
                        item.flag = !item.flag;
                    }
                    else
                    {
                        if (item.duration > 0)
                            key_down((int)key);
                        else
                            click((int)key);
                        if (item.duration > 0)
                        {
                            Thread.Sleep(item.duration);
                            key_up((int)key);
                        }
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
                    var savePath = Path.Combine(Application.StartupPath, "pics", $"{DateTime.Now.ToString("yyyMMddHHmmss")}.png");
                    ScreenCat.SaveImageWithQuality(savePath, screen, 70);

                    picCut.BeginInvoke(new Action(() => { picCut.Image = screen; }));

                    SetText($"截图成功:{savePath}");
                    Task.Run(() => { CheckPicture(savePath, CheckType.一般测谎); });
                    Task.Run(() => { CheckPicture(savePath, CheckType.解轮检测); });
                    Task.Run(() => { CheckPicture(savePath, CheckType.蘑菇测谎); });

                }
                else if (item.type == JobType.姑姑神社)
                {

                    var isFind = FindText(item);
                    if (!isFind) click((int)KeyBoard.Escape);
                    while (!isFind && isStaring)
                    {
                        //没有找到执行前面两个任务 
                        var idx = jobs.FindIndex(t => t == item);
                        StarJob(jobs[idx - 2]);
                        StarJob(jobs[idx - 1]);
                        isFind = FindText(item);
                        if (!isFind) click((int)KeyBoard.Escape);
                    }
                    //找到了
                    SetText("找到了姑姑"); 
                    click((int)KeyBoard.空格键);
                    Thread.Sleep(3000);
                    click((int)KeyBoard.空格键);
                    Thread.Sleep(3000);
                    click((int)KeyBoard.空格键);
                    Thread.Sleep(3000);
                    click((int)KeyBoard.Escape);

                    //切换任务
                    StopTask();
                    comTask.BeginInvoke(new Action(() => {
                        int tidx = 0;
                        foreach (var task in comTask.Properties.Items)
                        {
                            var tt = (ImageComboBoxItem)task;
                            if (tt.Description.Equals("011-姑姑神社领奖")) break;
                            tidx++;
                        }
                        comTask.SelectedIndex = tidx;
                        Task.Run(() =>
                        {
                            Thread.Sleep(1000);
                            //重置时间
                            foreach (var j in jobs)
                            { 
                                j.less = j.time - 1;
                            }
                            Start();
                        });
                    }));
                    return; 
                }
                else if (item.type == JobType.姑姑领奖)
                {
                    StopTask();
                    comTask.BeginInvoke(new Action(() => {
                        int tidx = 0;
                        foreach (var task in comTask.Properties.Items)
                        {
                            var tt = (ImageComboBoxItem)task;
                            if (tt.Description.Equals("010-姑姑神社开始")) break;
                            tidx++;
                        }
                        comTask.SelectedIndex = tidx;
                        Task.Run(() =>
                        {
                            Thread.Sleep(1000);
                            //重置时间
                            foreach (var j in jobs)
                            {
                                j.less = j.time;
                            }
                            Start();
                        }); 
                    }));
                    return;
                }

                if (item.less == 0)
                    item.less = item.time;

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

        private bool FindText(Job item)
        {
            var screen = ScreenCat.GetScreen(false, item.x, item.y, item.width, item.height);
            var savePath = Path.Combine(Application.StartupPath, "pics", $"{DateTime.Now.ToString("yyyMMddHHmmss")}.png");
            ScreenCat.SaveImageWithQuality(savePath, screen, 70);


            ////picCut.BeginInvoke(new Action(() => { picCut.Image = screen; }));

            //TesseractEngine ocr = new TesseractEngine("./tessdata", "chi_tra");//设置语言   繁体 

            //Bitmap bit = new Bitmap(Image.FromFile("D:\\hudingwen\\github\\my-OpenCV-Python\\pic\\mxd\\test\\sampleGuGu\\1.jpg"));
            ////bit = ToGray(bit);
            //picCut.BeginInvoke(new Action(() => { picCut.Image = bit; }));
            ////bit = PreprocesImage(bit);//进行图像处理,如果识别率低可试试
            //Page page = ocr.Process(bit, PageSegMode.SingleLine);
            //string str = page.GetText();//识别后的内容
            //page.Dispose();
            //SetText(str);


            Process p = new Process();
            string path = Path.Combine(Application.StartupPath, "py", "045-find_text.py"); ;//待处理python文件的路径，本例中放在debug文件夹下
            string sArguments = path;
            ArrayList arrayList = new ArrayList();
            //arrayList.Add("D:\\hudingwen\\github\\my-OpenCV-Python\\pic\\mxd\\test\\sampleGuGu\\4.jpg");
            arrayList.Add(savePath);
            foreach (var param in arrayList)//添加参数
            {
                sArguments += " " + param;
            }

            p.StartInfo.FileName = @"Python"; //python2.7的安装路径
            p.StartInfo.Arguments = sArguments;//python命令的参数
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardInput = true;
            p.StartInfo.RedirectStandardError = true;
            p.StartInfo.CreateNoWindow = true;
            p.Start();//启动进程 

            var msg_success = p.StandardOutput.ReadToEnd();  // 接收信息 
            var msg_error = p.StandardError.ReadToEnd(); //接收错误信息
            if (string.IsNullOrWhiteSpace(msg_success))
            {
                //没有正常返回
                SetText($"检测出错-查找文字-{msg_error}");
            }
            else
            {
                SetText($"查找文字-{msg_success}");
            }
            if(msg_success.Contains("採集藥草") || msg_success.Contains("砍樹"))
            {
                return true;
            }
            else
            {
                return false;
            }
            //採集藥草 砍樹
        }

        /// <summary>
        /// 图像灰度化算法（加权）
        /// </summary>
        /// <param name="inputBitmap">输入图像</param>
        /// <param name="outputBitmap">输出图像</param>
        public static Bitmap ToGray(Bitmap inputBitmap)
        {
            for (int i = 0; i < inputBitmap.Width; i++)
            {
                for (int j = 0; j < inputBitmap.Height; j++)
                {
                    Color color = inputBitmap.GetPixel(i, j);  //在此获取指定的像素的颜色
                    int gray = (int)(color.R * 0.3 + color.G * 0.59 + color.B * 0.11);
                    Color newColor = Color.FromArgb(gray, gray, gray);
                    //设置指定像素的颜色  参数：坐标x,坐标y，颜色
                    inputBitmap.SetPixel(i, j, newColor);
                }
            }
           return inputBitmap;
        }

        public void DeleHistoryPics()
        {
            try
            {
                while (isStaring)
                {
                    //间隔一段时间删除
                    var dic = Path.Combine(Application.StartupPath, "pics");
                    if (!Directory.Exists(dic)) Directory.CreateDirectory(dic);
                    var files = Directory.GetFiles(dic, "*.png");
                    if (files != null)
                   {
                        //删除前一百张
                        int length = files.Length - 100;
                        for (int i = 0; i < length; i++)
                        {
                            File.Delete(files[i]);
                        }
                    }
                    Thread.Sleep(30 * 60 * 1000);
                }
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
                                if (!isStaring) return;
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
                    for (int i = 0; i < jobs.Count; i++)
                    {
                        Job item = jobs[i];
                        if (item.less == 0) continue;
                        item.less = item.less - 15;
                        if (item.less < 0) item.less = 0;
                    }
                    Thread.Sleep(10);
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
            sp.Stop();
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
            var jtype = (JobType)((ImageComboBoxItem)this.comJob.SelectedItem).Value;
            if (jtype == JobType.按键任务)
            {
                job.content = (KeyBoard)((ImageComboBoxItem)this.comBoard.SelectedItem).Value;
            }
            else if (jtype == JobType.点击任务)
            {
                job.content = (EnumMouse)((ImageComboBoxItem)this.comMouse.SelectedItem).Value;
                job.x = (int)posX.Value;
                job.y = (int)posY.Value;
            }
            else if (jtype == JobType.聚焦窗体)
            {
                job.content = this.windowName.Text;
            }
            else if (jtype == JobType.移动窗体)
            {
                job.content = this.windowName.Text;
                job.x = Convert.ToInt32(posX.Value);
                job.y = Convert.ToInt32(posY.Value);
            }
            else if (jtype == JobType.截图检测)
            {
                job.content = this.windowName.Text;
            }else if(jtype== JobType.姑姑神社)
            {
                job.content = "";
                job.x = (int)posX.Value;
                job.y = (int)posY.Value;
                job.width = (int)numWidth.Value;
                job.height = (int)numHeight.Value;
            }
            else if (jtype == JobType.姑姑领奖)
            {
                job.content = "";
            }
            job.isAsync = checkAysnc.Checked;
            job.type = jtype;
            job.contentName = job.content.ToString();
            job.typeName = job.type.ToString();
            job.count = Convert.ToInt32(numCount.Value);
            job.remark = txtRemark.Text;
            job.enable = true;
        }

        private void SetTime(Job job)
        {
            job.time = (int)(Convert.ToDecimal(numTime.Value) * 1000);
            job.less = (int)(Convert.ToDecimal(numTime.Value) * 1000);
            job.delay = (int)(Convert.ToDecimal(numDelay.Text) * 1000);
            job.duration = (int)(Convert.ToDecimal(textDuration.Value) * 1000);
            job.isFirstRun = true;
        }

        public void SaveConfig()
        {
            //保存任务
            var config = Newtonsoft.Json.JsonConvert.SerializeObject(jobs, Formatting.Indented);
            File.WriteAllText(((ImageComboBoxItem)comTask.SelectedItem).Value.ToString(), config);

            ConfigInfo configInfo = new ConfigInfo();
            configInfo.isPlay = checkMusic.Checked;
            configInfo.playName = comMusic.Text;
            configInfo.taskName = comTask.Text;



            var settings = Newtonsoft.Json.JsonConvert.SerializeObject(configInfo, Formatting.Indented);
            var dic = Path.Combine(Application.StartupPath, "config");
            if (!Directory.Exists(dic)) Directory.CreateDirectory(dic);
            var setPath = Path.Combine(dic, "settings.config");

            bool inUse = true;
            //true表示正在使用,false没有使用
            FileStream fs = null;
            try
            {

                fs = new FileStream(setPath, FileMode.Open, FileAccess.Read,

                FileShare.None);

                inUse = false;
            }
            catch
            {

            }
            finally
            {
                if (fs != null)

                    fs.Close();
            }
            if (!inUse)
            {
                File.WriteAllText(setPath, settings);
            }


            

        }
        public void ReadConfig()
        {
            jobs.Clear();
            var filePath = ((ImageComboBoxItem)comTask.SelectedItem).Value.ToString();
            if (File.Exists(filePath))
            {
                var config = File.ReadAllText(filePath);
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
                var msg = ""; 
                mFocus.FocusWindow(windowName.Text, ref msg);
                var screen = ScreenCat.GetScreen(windowName.Text);
                var savePath = Path.Combine(Application.StartupPath, "pics", $"{DateTime.Now.ToString("yyyMMddHHmmss")}.png");
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
                
                

                ReadTask(jobName.Text);
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
                var filePath = ((ImageComboBoxItem)comTask.SelectedItem).Value.ToString();
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
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
                        renderer.WriteToStream(qrCode.Matrix, System.Drawing.Imaging.ImageFormat.Png, ms);
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
                if (!isStaring) return;
                sp.Stop();
                SetText("停止播放");
            }
            catch (Exception ex)
            {
                SetText(ex.Message);
            }
        }

        private void btnEnable_Click(object sender, EventArgs e)
        { 
            var row = viewJobs.GetSelectedRows();
            if (row.Length > 0)
            { 
                foreach (int idx in row)
                {
                    var tjob = (Job)viewJobs.GetRow(idx);
                    tjob.enable = !tjob.enable;
                }  
                RefreshData();
                SaveConfig();
                SetText("操作成功");
            }
            else
            {
                SetText("请选择要操作的数据");
            }
        }
    }
}
