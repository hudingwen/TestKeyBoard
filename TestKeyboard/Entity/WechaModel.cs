using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestKeyboard.Entity
{
    internal class WechatModel
    {
        public string status { get; set; }
        public bool success { get; set; }
        public string msg { get; set; }
        public string msgDev { get; set; }
        public WechatResponse response { get; set; }
       

    }
    public class WechatResponse
    {
        public string id { get; set; }
        public string companyCode { get; set; }
        public UserData usersData { get; set; }
    }

    public class UserData
    {
        public string id { get; set; }
        public string errcode { get; set; }
        public string errmsg { get; set; }
        public string access_token { get; set; }
        public string expires_in { get; set; }
        public string total { get; set; }
        public string count { get; set; }
        public string data { get; set; }
        public string users { get; set; }
        public string next_openid { get; set; }
        public string template_list { get; set; }
        public string menu { get; set; }
        public string ticket { get; set; }
        public string expire_seconds { get; set; }
        public string url { get; set; }
        public string subscribe { get; set; }
        public string openid { get; set; }
        public string nickname { get; set; }
        public string sex { get; set; }
        public string language { get; set; }
        public string city { get; set; }
        public string province { get; set; }
        public string country { get; set; }
        public string headimgurl { get; set; }
    }
}
