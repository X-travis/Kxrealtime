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

        public delegate void normalCb(string msg);
        public event normalCb MessageReceived;

        public WebSocketState State
        {
            get
            {
                return curWebScoket.State;
            }
        }


        public WebSocket webSocketInstance
        {
            get
            {
                return curWebScoket;
            }
        }

        public webSocketClient() { }

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

        private void CurWebScoket_MessageReceived(object sender, MessageReceivedEventArgs e)
        {
            MessageReceived(e.Message);
        }

        public void clientSend(string info)
        {
            Task.Run(() => curWebScoket.Send(info));
        }

        public void closeFn()
        {
            
        }

        public void websocket_Opened(object sender, EventArgs e)
        {
            
        }

        public void websocket_Closed(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("websocket close");
        }

        public void websocket_Error(object sender, SuperSocket.ClientEngine.ErrorEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("websocket close");
            reconnect();
        }

        private void reconnect()
        {
            StartWebSocket(curUrl);
        }

        public void Dispose()
        {
            curWebScoket.Close();
            curWebScoket.Dispose();
        }
    }
}
