using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace TokuchoBugyoK2
{
    public class ButtonTextCustomizableMessageBox
    {
        private IntPtr hHook = IntPtr.Zero;

        /// &lt;summary>
        /// ボタンに表示するテキストを指定します。
        /// &lt;/summary>
        public CustomButtonText ButtonText { get; set; }

        /// &lt;summary>
        /// コンストラクタ。
        /// &lt;/summary>
        public ButtonTextCustomizableMessageBox()
        {
            this.ButtonText = new CustomButtonText();
        }

        /// &lt;summary>
        /// ダイアログボックスを表示します。
        /// &lt;/summary>
        public DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icons)
        {
            try
            {
                BeginHook();
                return MessageBox.Show(text, caption, buttons, icons);
            }
            finally
            {
                EndHook();
            }
        }

        /// &lt;summary>
        /// フックを開始します。
        /// &lt;/summary>
        void BeginHook()
        {
            EndHook();
            this.hHook = SetWindowsHookEx(WH_CBT, new HOOKPROC(this.HookProc), IntPtr.Zero, GetCurrentThreadId());
        }

        IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode == HCBT_ACTIVATE)
            {
                if (this.ButtonText.Abort != null) SetDlgItemText(wParam, ID_BUT_ABORT, this.ButtonText.Abort);
                if (this.ButtonText.Cancel != null) SetDlgItemText(wParam, ID_BUT_CANCEL, this.ButtonText.Cancel);
                if (this.ButtonText.Ignore != null) SetDlgItemText(wParam, ID_BUT_IGNORE, this.ButtonText.Ignore);
                if (this.ButtonText.No != null) SetDlgItemText(wParam, ID_BUT_NO, this.ButtonText.No);
                if (this.ButtonText.OK != null) SetDlgItemText(wParam, ID_BUT_OK, this.ButtonText.OK);
                if (this.ButtonText.Retry != null) SetDlgItemText(wParam, ID_BUT_RETRY, this.ButtonText.Retry);
                if (this.ButtonText.Yes != null) SetDlgItemText(wParam, ID_BUT_YES, this.ButtonText.Yes);

                EndHook();
            }

            return CallNextHookEx(this.hHook, nCode, wParam, lParam);
        }

        /// &lt;summary>
        /// フックを終了します。何回呼んでもOKです。
        /// &lt;/summary>
        void EndHook()
        {
            if (this.hHook != IntPtr.Zero)
            {
                UnhookWindowsHookEx(this.hHook);
                this.hHook = IntPtr.Zero;
            }
        }

        #region Win32API

        const int WH_CBT = 5;
        const int HCBT_ACTIVATE = 5;

        const int ID_BUT_OK = 1;
        const int ID_BUT_CANCEL = 2;
        const int ID_BUT_ABORT = 3;
        const int ID_BUT_RETRY = 4;
        const int ID_BUT_IGNORE = 5;
        const int ID_BUT_YES = 6;
        const int ID_BUT_NO = 7;

        private delegate IntPtr HOOKPROC(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(int idHook, HOOKPROC lpfn, IntPtr hInstance, IntPtr threadId);

        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(IntPtr hHook);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hHook, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetCurrentThreadId();

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool SetDlgItemText(IntPtr hWnd, int nIDDlgItem, string lpString);

        #endregion
    }
}
