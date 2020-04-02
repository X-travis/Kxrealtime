using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.WebSockets;
using System.Text;
using System.Threading.Tasks;
using Websocket.Client;

namespace kxrealtime.utils
{
    public static class webSocketClient
    {
        public static IWebsocketClient StartWebSocket(string uri)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("websocket init");
                //var exitEvent = new ManualResetEvent(false);
                var url = new Uri(uri);
                var factory = new Func<ClientWebSocket>(() =>
                {
                    var client = new ClientWebSocket
                    {
                        Options =
                        {
                            KeepAliveInterval = TimeSpan.FromSeconds(30),
                            // Proxy = ...
                            // ClientCertificates = ...
                        }
                    };
                    //client.Options.SetRequestHeader("Origin", "https://kx-v010.dev.resfair.com");
                    return client;
                });
                IWebsocketClient websocketClent = new WebsocketClient(url, factory);
                websocketClent.ReconnectTimeout = TimeSpan.FromSeconds(30);
                websocketClent.ReconnectionHappened.Subscribe(info => System.Diagnostics.Debug.WriteLine("reconnect " + info.Type));
                websocketClent.DisconnectionHappened.Subscribe(info =>
                    System.Diagnostics.Debug.WriteLine($"Disconnection happened, type: {info.Type}"));

                System.Diagnostics.Debug.WriteLine(websocketClent.Url);
                websocketClent.Start();
                return websocketClent;
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e);
                return null;
            }

            //Task.Run(() => client.Send("{ message }"));
        }

        public static void clientSend(IWebsocketClient client, string info)
        {
            //await client.Send("test");
            Task.Run(() => client.Send(info));
        }

        public static void closeFn()
        {

        }
    }
}
