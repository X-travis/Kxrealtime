using kxrealtime.kxdata;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace kxrealtime.utils
{
    public static class request
    {
        public delegate void ImgCb(string s);

        //使用方法之前获得cookie并放入RestClient对象
        public static RestClient GetClient()
        {
            System.Net.ServicePointManager.ServerCertificateValidationCallback += (s, cert, chain, sslPolicyErrors) => true;
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12 | System.Net.SecurityProtocolType.Tls11 | System.Net.SecurityProtocolType.Tls;

            //RestClient传递的string型的url
            var client = new RestClient();
            //client.CookieContainer = new System.Net.CookieContainer();
            //cookies.ForEach(c => client.CookieContainer.Add(new System.Net.Cookie(c.Name, c.Value, c.Path, c.Domain)));
            return client;
        }

        //封装了一个方法来给restClient调用
        public static IRestResponse SendRequest(this RestClient client, Uri requestUrl, Method method, Dictionary<string, string> parameters = null, Object jsonBody = null)
        {
            var request = client.FirmRequest(method);
            client.BaseUrl = requestUrl;
            if (parameters != null)
            {
                foreach (var p in parameters)
                {
                    request.AddParameter(p.Key, p.Value, ParameterType.QueryString);
                }
            }
            if (jsonBody != null)
            {
                request.AddJsonBody(jsonBody);
            }
            //调用Execute方法来执行request
            return client.Execute(request);
        }
        public static RestRequest FirmRequest(this RestClient client, Method method)
        {
            var request = new RestRequest();
            request.Method = method;
            request.AddHeader("Content-Type", "application/json");
            request.AddHeader("Accept", "application/json, text/plain, text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");

            return request;
        }

        // 上传图片
        public static void UploadImg(string postUrl, string file, int curIdx)
        {
            var task = new Task<upsertTeachContent>(() =>
            {
                try
                {
                    var client = new RestClient(postUrl);
                    var request = new RestRequest(Method.POST);
                    request.AddHeader("Cache-Control", "no-cache");
                    request.AddHeader("content-type", "multipart/form-data");
                    request.AddFile("file", file);
                    //request.AddParameter("pdfjson", postDataStr);
                    var response = client.Execute(request);
                    JObject data = JObject.Parse(response.Content);
                    string code = (string)data["code"];
                    string url = "";
                    if (code != "0")
                    {
                        System.Diagnostics.Debug.WriteLine("initchapter" + code);
                    }
                    else
                    {
                        try
                        {
                            url = (string)data["data"];
                            return recordTchImg(url);
                        }
                        catch (Exception e) { }

                    }
                    return null;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            });
            task.ContinueWith((Task<upsertTeachContent> result) =>
            {
                if(result == null || result.Result == null)
                {
                    return;
                }
                upsertTeachContentItem itemTmp = result.Result.teach_content_list[0];
                object oData = (new
                {
                    key = "classroom",
                    value = utils.KXINFO.KXCHOSECLASSID,
                    type = "PPT",
                    data = new
                    {
                        img = itemTmp.snapshot,
                        imgId = itemTmp.tid,
                        pn = curIdx
                    },
                    timestamp = utils.Utils.getTimeStamp()
                }) ;
                JObject o = JObject.FromObject(oData);
                string tmp = o.ToString();
                Globals.ThisAddIn.SendTchInfo(tmp);
                recordTch(oData);
            });
            task.Start();
        }

        public static upsertTeachContent recordTchImg(string url)
        {
            List<object> args = new List<object>
            {
                new
                {
                    teach_record_id = Int64.Parse(utils.KXINFO.KXTCHRECORDID),
                    snapshot = url
                }
            };
            var client = new RestClient($"{utils.KXINFO.KXURL}/usr/upsertTeachContent?session_id={utils.KXINFO.KXSID}");
            var request = new RestRequest(Method.POST);
            request.AddHeader("Cache-Control", "no-cache");
            request.AddJsonBody(args);
            //request.AddParameter("pdfjson", postDataStr);
            var response = client.Execute(request);
            if (response.ErrorException != null)
            {
                utils.Utils.LOG("upsertTeachContent response.ErrorException error: " + response.ErrorException.Message);
            } else
            {
                JObject data = JObject.Parse(response.Content);
                string code = (string)data["code"];
                if (code != "0")
                {
                    utils.Utils.LOG("upsertTeachContent request error: " + code);
                }
                else
                {
                    string dataTmp = data["data"].ToString();
                    upsertTeachContent curData = JsonConvert.DeserializeObject<upsertTeachContent>(dataTmp);
                    return curData;
                }
            }
            return null;
        }

        public static void recordTch(object stepContent)
        {
            Uri reqUrl = new Uri($"{utils.KXINFO.KXURL}/usr/upsertTeachRecord");
            Dictionary<string, string> args = new Dictionary<string, string> { };
            args.Add("session_id", utils.KXINFO.KXSID);
            List<object> recordArgsArr = new List<object> {
                new
                {
                    class_id = utils.KXINFO.KXCHOSECLASSID,
                    course_id = utils.KXINFO.KXCHOSECOURSEID,
                    chapter_id = utils.KXINFO.KXCHOSECHAPTERID,
                    step_content = new List<object>(){stepContent }
                }
            };
            RestSharp.IRestResponse response = utils.request.SendRequest(Globals.ThisAddIn.CurHttpReq, reqUrl, RestSharp.Method.POST, args, recordArgsArr);
            if (response.ErrorException != null)
            {
                utils.Utils.LOG("upsertTeachRecord response.ErrorException error: " + response.ErrorException.Message);
            }
            else
            {
                JObject data = JObject.Parse(response.Content);
                string code = (string)data["code"];
                if (code != "0")
                {
                    utils.Utils.LOG("upsertTeachRecord  error code: " + code);
                }
                else
                {

                }
            }
        }
    }
}
