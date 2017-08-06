using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace IndigoQRCode
{
    [Guid("35813A06-1C88-45B1-AE0E-D791602AEEDF")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IIndigoQRCode
    {
        [Description("値")]
        string Value { get; set; }
    }
}
