using System;
using System.Runtime.InteropServices;

namespace IndigoQRCode
{
    [Guid("54A51C91-5C40-49F4-A22A-254BEBDE18CA")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IIndigoQRCodeEvent
    {
        [DispId(1)]
        void ValueChanged(string oldValue);
    }
}
