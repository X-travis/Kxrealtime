using Newtonsoft.Json.Linq;
using System;

namespace kxrealtime.utils
{
    public static class KXINFO
    {

        public static string KXURL = @"https://kxrealtime-mp.kuxiao.cn";
        public static string KXADMINURL = @"https://kxrealtime-admin.kuxiao.cn";
        public static string KXSOCKETURL = @"wss://kxrealtime-mp.kuxiao.cn";
        public static string KXCOURSEURL = @"https://course.kuxiao.cn";

        //public static string KXEXAMSTAT = @"http://kxrealtime-wap.dev.gdy.io/#/pptComponents/countAnswerChart";

        public static int ChoseColor = 65280; // System.Drawing.Color.FromArgb(1, 0, 255, 0).ToArgb();
        // 用户信息相关
        public static string KXTOKEN;
        public static string KXUID;
        public static string KXSID;
        public static string KXOUTUID;
        public static string KXUNAME;
        public static string KXUAVATAR;

        // 授课相关
        public static string KXTCHRECORDID;
        public static Int64 KXCHOSECOURSEID;
        public static Int64 KXCHOSECLASSID;
        public static Int64 KXCHOSECHAPTERID;
        public static string KXCHOSECOURSETITLE;
        public static string KXCHOSECLASSNAME;
        public static string KXCHOSECHAPTERTITLE;

        public static long srvTimeDif = 0;


        public static void initUsr(string dataInfo)
        {
            JObject data = JObject.Parse(dataInfo);
            KXTOKEN = (string)data["kx_token"];
            KXSID = (string)data["session_id"];
            KXUID = (string)data["user"]["kuxiao_uid"];
            KXOUTUID = (string)data["user"]["tid"];
            try
            {
                KXUNAME = (string)data["kx_user_info"]["data"]["usr"]["attrs"]["basic"]["nickName"];
                KXUAVATAR = (string)data["kx_user_info"]["data"]["usr"]["attrs"]["basic"]["avatar"];
            }
            catch (Exception e)
            {
                utils.Utils.LOG("parse kx_user_info.data.usr.atts.basic error" + e.Message);
            }
            try
            {
                var srvTime = (long)data["server_time"];
                srvTimeDif = srvTime - utils.Utils.getTimeStamp();
            }
            catch (Exception) { }
        }

        public static void clear()
        {
            KXTOKEN = null;
            KXSID = null;
            KXUID = null;
            KXOUTUID = null;
            KXUNAME = null;
            KXUAVATAR = null;

            KXCHOSECHAPTERTITLE = null;
            KXTCHRECORDID = null;
            KXCHOSECLASSID = 0;
            KXCHOSECOURSEID = 0;
            KXCHOSECHAPTERID = 0;
            srvTimeDif = 0;
        }

        public static void tchClear()
        {
            KXCHOSECHAPTERTITLE = null;
            KXTCHRECORDID = null;
            KXCHOSECLASSID = 0;
            KXCHOSECOURSEID = 0;
            KXCHOSECHAPTERID = 0;
        }
    }
}
