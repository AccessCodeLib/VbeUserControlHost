using System;
using System.Runtime.InteropServices;


namespace AccessCodeLib.Common.VBIDETools
{
    [ComVisible(true)]
    [Guid("007610E2-E9A0-498E-867E-36AA8BC4D71C")]
    public interface IVbeUserControlHost : IVbeWindowsFormsUserControlHost, IVbeWPFUserControlHost
    {
    }

    [ComVisible(true)]
    [Guid("326CB39E-C621-41F6-833D-BACE54D26763")]
    public interface IVbeWindowsFormsUserControlHost
    {
        void HostUserControl(System.Windows.Forms.UserControl UserControlToHost);
    }

    [ComVisible(true)]
    [Guid("A5014EE3-82C7-4790-B12F-4EE2C95092D2")]
    public interface IVbeWPFUserControlHost
    {
        void HostUserControl(System.Windows.Controls.UserControl UserControlToHost);
    }


}
