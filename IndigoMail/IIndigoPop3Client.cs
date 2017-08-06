using System;
using System.ComponentModel;
using System.Runtime.InteropServices;


namespace IndigoMail
{
    [Guid("A70C8500-5F5C-4E0E-B54B-E5B963BCFE97")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IIndigoPop3Client
    {
        [Description("メールサーバーのホスト名")]
        string Host { get; set; }

        [Description("ポート番号")]
        int Port { get; set; }

        [Description("SSL を使用するか否か")]
        bool UseSsl { get; set; }

        [Description("認証します。")]
        void Authenticate(string userName, string password);
    }
}
