using MailKit.Net.Pop3;
using System;
using System.Runtime.InteropServices;

namespace IndigoMail
{
    [Guid("2B4C2143-D09B-4F13-B30D-E782FFCC5E3E")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IIndigoPop3Client))]
    public class IndigoPop3Client : IIndigoPop3Client
    {
        public string Host { get; set; }
        public int Port { get; set; }
        public bool UseSsl { get; set; }

        public void Authenticate(string userName, string password)
        {
            Execute(client =>
            {
                // 認証
                client.Authenticate(userName, password);
            });
        }

        #region プライベートメソッド

        void Execute(Action<Pop3Client> action)
        {
            using (var client = new Pop3Client())
            {
                // 接続
                client.Connect(this.Host, this.Port, this.UseSsl);

                try
                {
                    action(client);
                }
                finally
                {
                    // 接続されている場合
                    if (client.IsConnected)
                    {
                        client.Disconnect(true);
                    }
                }
            }
        }
        #endregion
    }
}
