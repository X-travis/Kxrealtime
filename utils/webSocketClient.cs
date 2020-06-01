using System;
using System.Threading.Tasks;
using UnityEngine;
using Websocket.Client;
using WebSocket4Net;

namespace kxrealtime.utils
{
    public class webSocketClient : IDisposable
    {
        private string curUrl;
        private WebSocket curWebScoket;
        // 代理函数
        public delegate void normalCb(string msg);
        // 收到数据的事件
        public event normalCb MessageReceived;

        // 暴露的当前websocket 状态
        public WebSocketState State
        {
            get
            {
                return curWebScoket.State;
            }
        }

        // 获取实例
        public WebSocket webSocketInstance
        {
            get
            {
                return curWebScoket;
            }
        }

        public webSocketClient() { }

        // webscoket链接
        public WebSocket StartWebSocket(string uri)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("websocket init");
                curUrl = uri;
                curWebScoket = new WebSocket(uri);
                curWebScoket.Opened += new EventHandler(websocket_Opened);
                curWebScoket.Error += websocket_Error;
                curWebScoket.Closed += new EventHandler(websocket_Closed);
                curWebScoket.MessageReceived += CurWebScoket_MessageReceived;
                //websocket.MessageReceived += new EventHandler(websocket_MessageReceived);
                curWebScoket.Open();
                return webSocketInstance;
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e);
                utils.Utils.LOG(e.Message);
                return null;
            }
        }

        // 收到数据回调
        private void CurWebScoket_MessageReceived(object sender, MessageReceivedEventArgs e)
        {
            MessageReceived(e.Message);
        }

        // 发送数据
        public void clientSend(string info)
        {
            Task.Run(() => curWebScoket.Send(info));
        }

        public void closeFn()
        {
            
        }

        // 打开回调
        public void websocket_Opened(object sender, EventArgs e)
        {
            
        }

        // 关闭回调
        public void websocket_Closed(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("websocket close");
        }

        public void websocket_Error(object sender, SuperSocket.ClientEngine.ErrorEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("websocket close");
            reconnect();
        }

        // 重连方法
        private void reconnect()
        {
            StartWebSocket(curUrl);
        }

        // 销毁方法
        public void Dispose()
        {
            curWebScoket.Close();
            curWebScoket.Dispose();
        }
    }
}
