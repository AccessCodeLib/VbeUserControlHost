using System;
using System.Drawing;
using System.Runtime.InteropServices;
using Forms = System.Windows.Forms;
using WPF = System.Windows.Controls;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.Common.VBIDETools
{
    public class VbeUserControl<TControl> : IDisposable
    {
        private readonly TControl _control;
        private Window _vbeWindow;

        public VbeUserControl(AddIn addIn, string caption, string positionGuid, 
                                TControl controlToHost, bool visible = true,
                                string vbideUserControlHostProgId = VbeUserControlHostSettings.ProgId)
        {
            object docObj = null;
            _vbeWindow = addIn.VBE.Windows.CreateToolWindow(addIn, vbideUserControlHostProgId,
                                                            caption, positionGuid, ref docObj);
            _vbeWindow.Visible = true;

            _control = controlToHost;
            if (_control is Forms.UserControl winFormControl)
            {
                if (!(docObj is IVbeWindowsFormsUserControlHost userControlHost))
                {
                    throw new InvalidComObjectException(string.Format("docObj cannot be casted to IVbeWindowsFormsUserControlHost"));
                }
                userControlHost.HostUserControl(winFormControl);
            }
            else if (_control is WPF.UserControl wpfControl)
            {
                if (!(docObj is IVbeWPFUserControlHost userControlHost))
                {
                    throw new InvalidComObjectException(string.Format("docObj cannot be casted to IVbeWPFUserControlHost"));
                }
                userControlHost.HostUserControl(wpfControl);
            }
            else
            {
                throw new ArgumentException("controlToHost must be a System.Windows.Forms.UserControl or a System.Windows.Controls.UserControl");
            }       
            
            if (!visible)
            {
                _vbeWindow.Visible = false;
            }
        }

        public TControl Control { get { return _control; } }
        private Window VbeWindow { get { return _vbeWindow; } }

        public bool Visible
        {
            get
            {
                try
                {
                    return VbeWindow.Visible;
                }
                catch (Exception)
                {
                    return false;
                }
            }
            set { VbeWindow.Visible = value; }
        }

        public void Show()
        {
            if (!Visible)
            {
                Visible = true;
            }   
        }

        #region IDisposable Support

        bool _disposed;

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                DisposeManagedResources();
            }
            DisposeUnmanagedResources();
            _disposed = true;
        }

        private void DisposeManagedResources()
        {
            if (_control is IDisposable disposableControl)
            {
                disposableControl.Dispose();
            }
        }   

        private void DisposeUnmanagedResources()
        {
            if (_vbeWindow != null)
            {
                try
                {
                    _vbeWindow.Close();
                }
                catch { /* ignore */ }
                finally
                {
                    Marshal.ReleaseComObject(_vbeWindow);
                    _vbeWindow = null;
                }
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~VbeUserControl()
        {
            Dispose(false);
        }
        #endregion
    }

    internal class SubClassingResizeWindow : Forms.NativeWindow
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [StructLayout(LayoutKind.Sequential)]
        struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private readonly Forms.Control _userControl;
        private readonly Window _vbeWindow;
        private Size _lastSize;

        public SubClassingResizeWindow(Forms.Control userControl, Window vbeWindow)
        {
            _userControl = userControl;
            _vbeWindow = vbeWindow;

            IntPtr hWnd = FindVbeWindowHostHwnd(_vbeWindow);
            if (hWnd == IntPtr.Zero)
                throw new Exception(string.Concat("hWnd for ", _vbeWindow.Caption, " not found"));

            base.AssignHandle(hWnd);

            CheckSize();
        }

        private static IntPtr FindVbeWindowHostHwnd(Window vbeWindow)
        {
            const string DockedWindowClass = "wndclass_desked_gsk";
            const string FloatingWindowClass = "VBFloatingPalette";
            const string GenericPaneClass = "GenericPane";

            IntPtr hWnd;
            if (IsDockedWindow(vbeWindow))
            {
                hWnd = FindWindow(DockedWindowClass, vbeWindow.LinkedWindowFrame.Caption);
            }
            else
            {
                hWnd = FindWindow(FloatingWindowClass, vbeWindow.Caption);
            }
            hWnd = FindWindowEx(hWnd, IntPtr.Zero, GenericPaneClass, vbeWindow.Caption);
            return hWnd;
        }

        private static bool IsDockedWindow(Window vbeWindow)
        {
            return vbeWindow.LinkedWindowFrame.Type == vbext_WindowType.vbext_wt_MainWindow;
        }

        private void CheckSize()
        {
            GetClientRect(Handle, out RECT rect);
            var newSize = new Size(rect.Right - rect.Left, rect.Bottom - rect.Top);
            if (newSize != _lastSize)
            {
                _userControl.Width = newSize.Width;
                _userControl.Height = newSize.Height;
                _lastSize = newSize;
            }
        }

        protected override void WndProc(ref Forms.Message m)
        {
            const int WM_SIZE = 0x0005;

            if (m.Msg == WM_SIZE)
            {
                CheckSize();
            }
            base.WndProc(ref m);
        }
    }
}
