using Microsoft.Win32;
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ZXing;

namespace IndigoQRCode
{
    [Guid("80B3CDD4-02C5-4ED6-9995-A02B6CF06E93")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IIndigoQRCode))]
    [ComSourceInterfaces(typeof(IIndigoQRCodeEvent))]
    public partial class IndigoQRCode : UserControl, IIndigoQRCode
    {
        #region Events
        [ComVisible(false)]
        public delegate void ValueChangedEventHandler(string oldValue);
        public event ValueChangedEventHandler ValueChanged;
        #endregion

        string _value = string.Empty;

        public IndigoQRCode()
        {
            InitializeComponent();
        }

        public string Value
        {
            get
            {
                return _value;
            }
            set
            {
                SetValue(value);
            }
        }

        #region

        void SetValue(string contents)
        {
            string oldValue = _value;
            _value = contents;

            SetImage(contents);

            if (ValueChanged != null)
            {
                ValueChanged(oldValue);
            }
        }

        void SetImage(string contents)
        {
            var writer = new BarcodeWriter();
            // QRコード
            writer.Format = BarcodeFormat.QR_CODE;

            var image = writer.Write(contents);

            pictureBox1.Image = image;
        }
        #endregion

        #region レジストリー登録
        [ComRegisterFunction]
        static void Register(Type t)
        {
            var keyName = GetKeyName(t);
            // このライブラリのパス
            var fileName = Assembly.GetExecutingAssembly().Location;
            // グローバル属性情報を取得
            var ada = Attribute.GetCustomAttribute(
                Assembly.GetExecutingAssembly(),
                typeof(AssemblyDescriptionAttribute)
                ) as AssemblyDescriptionAttribute;

            // ActiveX コントロールのレジストリ情報を登録
            using (var key = Registry.ClassesRoot.OpenSubKey(keyName, true))
            {
                // 説明が設定されている場合
                if (!string.IsNullOrWhiteSpace(ada.Description))
                {
                    key.SetValue(string.Empty, ada.Description);
                }

                using (var subKey = key.CreateSubKey("DefaultIcon"))
                {
                    subKey.SetValue(string.Empty, $"{fileName},101");
                }

                using (var subKey = key.CreateSubKey("ToolboxBitmap32"))
                {
                    subKey.SetValue(string.Empty, $"{fileName},102");
                }

                key.CreateSubKey("Control").Close();

                using (var subKey = key.CreateSubKey("MiscStatus"))
                {
                    var value = (long)(OLEMISC.OLEMISC_INSIDEOUT
                        | OLEMISC.OLEMISC_ACTIVATEWHENVISIBLE
                        | OLEMISC.OLEMISC_SETCLIENTSITEFIRST);

                    subKey.SetValue(string.Empty, value);
                }

                using (var subKey = key.CreateSubKey("TypeLib"))
                {
                    var guid = Marshal.GetTypeLibGuidForAssembly(t.Assembly);
                    subKey.SetValue(string.Empty, guid.ToString("B"));
                }

                using (var subKey = key.CreateSubKey("Version"))
                {
                    var version = t.Assembly.GetName().Version;
                    subKey.SetValue(string.Empty, $"{version.Major}.{version.Minor}");
                }
            }
        }

        [ComUnregisterFunction]
        static void Unregister(Type t)
        {
            // 何もしない
        }

        static string GetKeyName(Type t)
        {
            return $"CLSID\\{t.GUID.ToString("B")}";
        }
        #endregion
    }
}
