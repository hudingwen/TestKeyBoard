using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestKeyboard.Entity
{
    public class job
    {
        /// <summary>
        /// 任务间隔时间(秒)
        /// </summary>
        public decimal time { get; set; }
        /// <summary>
        /// 剩余时间(秒)
        /// </summary>
        public decimal less { get; set; }
        /// <summary>
        /// 任务类型
        /// </summary>
        public jobType type { get; set; }
        public string typeName { get; set; }
        /// <summary>
        /// 任务内容
        /// </summary>
        public object content { get; set; }
        /// <summary>
        /// 点击x坐标
        /// </summary>
        public int x { get; set; }
        /// <summary>
        /// 点击y坐标
        /// </summary>
        public int y { get; set; }  
        /// <summary>
        /// 后置延迟
        /// </summary>
        public decimal delay { get; set; }
        public string contentName { get; set; }
        /// <summary>
        /// 运行次数
        /// </summary>
        public int count { get; set; }
        /// <summary>
        /// 已经运行次数
        /// </summary>
        public int countLess { get; set; }
    }
}
